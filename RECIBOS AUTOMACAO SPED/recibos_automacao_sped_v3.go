package main

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/csv"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"math"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

const docsBase = "https://docs.iob.com.br/api"

var transientStatusCodes = map[int]struct{}{
	http.StatusBadGateway:         {},
	http.StatusServiceUnavailable: {},
	http.StatusGatewayTimeout:     {},
	http.StatusTooManyRequests:    {},
}

type Config struct {
	HeadersPath     string
	OutputDir       string
	Period          string
	Workers         int
	RequestsPerSec  float64
	RequestTimeout  time.Duration
	ListRetries     int
	DetailRetries   int
	DownloadRetries int
	SkipExisting    bool
}

type Logger struct {
	mu   sync.Mutex
	file *os.File
}

type LogContext struct {
	Group    string
	Workflow string
	Empresa  string
}

type FailureEntry struct {
	GroupKey   string `json:"group_key"`
	Empresa    string `json:"empresa"`
	CNPJ       string `json:"cnpj"`
	Periodo    string `json:"periodo"`
	WorkflowID string `json:"workflow_id"`
	DocumentID string `json:"document_id"`
	Document   string `json:"document_name"`
	OutputPath string `json:"output_path"`
	Reason     string `json:"reason"`
}

type FailureCollector struct {
	mu      sync.Mutex
	entries []FailureEntry
}

type APIClient struct {
	baseURL string
	client  *http.Client
	headers map[string]string
	limiter *RateLimiter
	logger  *Logger
}

type RateLimiter struct {
	ch   <-chan time.Time
	stop func()
}

type WorkflowStub struct {
	ID           string
	State        int
	Category     string
	DocType      string
	Company      string
	CNPJ         string
	Period       string
	EndDate      time.Time
	CreationDate time.Time
	Raw          map[string]any
}

type WorkflowGroup struct {
	Key       string
	Company   string
	CNPJ      string
	Period    string
	Category  string
	DocType   string
	Workflows []WorkflowStub
}

type DocumentCandidate struct {
	ID        string
	Name      string
	Extension string
	MimeType  string
	Score     int
	MatchKind string
}

type DownloadResult struct {
	Path         string
	DocumentID   string
	DocumentName string
}

type ProcessSummary struct {
	Success  int
	Failures int
}

type jwtClaims struct {
	IAT int64 `json:"iat"`
	EXP int64 `json:"exp"`
}

func main() {
	cfg := defaultConfig()
	flag.StringVar(&cfg.HeadersPath, "headers", cfg.HeadersPath, "Caminho para headers_iob.json")
	flag.StringVar(&cfg.OutputDir, "out", cfg.OutputDir, "Diretório de saída dos PDFs")
	flag.StringVar(&cfg.Period, "period", cfg.Period, "Período alvo no formato MMyyyy")
	flag.IntVar(&cfg.Workers, "workers", cfg.Workers, "Quantidade de grupos processados em paralelo")
	flag.Float64Var(&cfg.RequestsPerSec, "rps", cfg.RequestsPerSec, "Máximo de requisições por segundo")
	flag.DurationVar(&cfg.RequestTimeout, "timeout", cfg.RequestTimeout, "Timeout por requisição")
	flag.IntVar(&cfg.ListRetries, "list-retries", cfg.ListRetries, "Tentativas para listar workflows")
	flag.IntVar(&cfg.DetailRetries, "detail-retries", cfg.DetailRetries, "Tentativas para fullDetails")
	flag.IntVar(&cfg.DownloadRetries, "download-retries", cfg.DownloadRetries, "Tentativas para download do PDF")
	flag.BoolVar(&cfg.SkipExisting, "skip-existing", cfg.SkipExisting, "Pula arquivo já existente no disco")
	flag.Parse()

	if err := os.MkdirAll(cfg.OutputDir, 0o755); err != nil {
		fatalf("não foi possível criar diretório de saída: %v", err)
	}

	logPath := filepath.Join(cfg.OutputDir, "log_go.txt")
	logger, err := NewLogger(logPath)
	if err != nil {
		fatalf("não foi possível abrir log: %v", err)
	}
	defer logger.Close()

	logger.Log(LogContext{}, strings.Repeat("=", 80))
	logger.Log(LogContext{}, "Sessão: %s", timestamp())
	logger.Log(LogContext{}, strings.Repeat("=", 80))

	headersPath, err := resolveHeadersPath(cfg.HeadersPath)
	if err != nil {
		fatalf("erro ao localizar headers: %v", err)
	}
	headers, err := loadHeaders(headersPath)
	if err != nil {
		fatalf("erro ao carregar headers: %v", err)
	}

	logger.Log(LogContext{}, "Headers carregados do arquivo JSON: %s", headersPath)
	logTokenExpirations(logger, headers)
	if err := abortIfTokensExpired(headers, 60*time.Second); err != nil {
		fatalf("%v", err)
	}

	apiClient := NewAPIClient(cfg, headers, logger)
	defer apiClient.Close()

	logger.Log(LogContext{}, "Período alvo (metadata.fileMetadata.period) = %s", cfg.Period)

	workflowItems, err := fetchWorkflowList(context.Background(), apiClient, cfg.ListRetries, logger)
	if err != nil {
		fatalf("erro ao listar workflows: %v", err)
	}
	if len(workflowItems) == 0 {
		logger.Log(LogContext{}, "Nenhum workflow retornado pela API. Encerrando.")
		return
	}

	groups := buildWorkflowGroups(workflowItems, cfg.Period, logger)
	if len(groups) == 0 {
		logger.Log(LogContext{}, "Nenhum workflow elegível após filtros. Encerrando.")
		return
	}

	logger.Log(LogContext{}, "Total de grupos após deduplicação: %d", len(groups))

	failures := &FailureCollector{}
	results := processGroups(context.Background(), cfg, apiClient, logger, failures, groups)

	saveFailures(cfg.OutputDir, logger, failures)

	logger.Log(LogContext{}, "Processamento finalizado. PDFs ok=%d, falhas=%d", results.Success, results.Failures)
	logger.Log(LogContext{}, "Log salvo em: %s", logPath)
}

func defaultConfig() Config {
	home, _ := os.UserHomeDir()
	out := filepath.Join(home, "Downloads", "RECIBOS_SPED")
	return Config{
		HeadersPath:     "headers_iob.json",
		OutputDir:       out,
		Period:          previousMonthPeriod(time.Now()),
		Workers:         6,
		RequestsPerSec:  5.0,
		RequestTimeout:  25 * time.Second,
		ListRetries:     3,
		DetailRetries:   4,
		DownloadRetries: 6,
		SkipExisting:    true,
	}
}

func fatalf(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}

func previousMonthPeriod(now time.Time) string {
	first := time.Date(now.Year(), now.Month(), 1, 0, 0, 0, 0, now.Location())
	prev := first.AddDate(0, 0, -1)
	return fmt.Sprintf("%02d%04d", int(prev.Month()), prev.Year())
}

func timestamp() string {
	return time.Now().Format("2006-01-02 15:04:05.000")
}

func NewLogger(path string) (*Logger, error) {
	file, err := os.OpenFile(path, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0o644)
	if err != nil {
		return nil, err
	}
	return &Logger{file: file}, nil
}

func (l *Logger) Close() error {
	if l == nil || l.file == nil {
		return nil
	}
	return l.file.Close()
}

func (l *Logger) Log(ctx LogContext, format string, args ...any) {
	line := fmt.Sprintf(format, args...)
	full := fmt.Sprintf("[%s][%s][%s](%s) %s", timestamp(), ctx.Group, ctx.Workflow, ctx.Empresa, line)

	l.mu.Lock()
	defer l.mu.Unlock()

	fmt.Println(full)
	_, _ = l.file.WriteString(full + "\n")
}

func loadHeaders(path string) (map[string]string, error) {
	raw, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var data map[string]any
	if err := json.Unmarshal(raw, &data); err != nil {
		return nil, err
	}

	headers := make(map[string]string, len(data))
	for key, value := range data {
		headers[key] = fmt.Sprint(value)
	}
	deleteHeader(headers, "if-none-match")
	deleteHeader(headers, "if-modified-since")
	return headers, nil
}

func resolveHeadersPath(input string) (string, error) {
	candidates := []string{
		input,
		filepath.Join(".", input),
		filepath.Join(".", "RECIBOS AUTOMACAO SPED", input),
	}
	for _, candidate := range candidates {
		if candidate == "" {
			continue
		}
		if info, err := os.Stat(candidate); err == nil && !info.IsDir() {
			return candidate, nil
		}
	}
	return "", fmt.Errorf("arquivo de headers não encontrado a partir de %q", input)
}

func deleteHeader(headers map[string]string, target string) {
	for key := range headers {
		if strings.EqualFold(key, target) {
			delete(headers, key)
		}
	}
}

func logTokenExpirations(logger *Logger, headers map[string]string) {
	if auth := firstHeader(headers, "authorization"); auth != "" && strings.HasPrefix(strings.ToLower(auth), "bearer ") {
		if claims, err := decodeJWT(strings.TrimSpace(auth[7:])); err == nil && claims.EXP > 0 && claims.IAT > 0 {
			iat := time.Unix(claims.IAT, 0)
			exp := time.Unix(claims.EXP, 0)
			logger.Log(LogContext{}, "Token 'authorization' iat=%s, exp=%s (duração ~%.1f minutos).",
				iat.Format("2006-01-02 15:04:05"),
				exp.Format("2006-01-02 15:04:05"),
				exp.Sub(iat).Minutes(),
			)
		}
	}

	if token := firstHeader(headers, "x-hypercube-idp-access-token"); token != "" {
		if claims, err := decodeJWT(token); err == nil && claims.EXP > 0 && claims.IAT > 0 {
			iat := time.Unix(claims.IAT, 0)
			exp := time.Unix(claims.EXP, 0)
			logger.Log(LogContext{}, "Token 'x-hypercube-idp-access-token' iat=%s, exp=%s (duração ~%.1f minutos).",
				iat.Format("2006-01-02 15:04:05"),
				exp.Format("2006-01-02 15:04:05"),
				exp.Sub(iat).Minutes(),
			)
		}
	}
}

func abortIfTokensExpired(headers map[string]string, skew time.Duration) error {
	var expired []string
	if auth := firstHeader(headers, "authorization"); auth != "" && strings.HasPrefix(strings.ToLower(auth), "bearer ") {
		claims, err := decodeJWT(strings.TrimSpace(auth[7:]))
		if err != nil || tokenExpired(claims, skew) {
			expired = append(expired, "authorization")
		}
	}
	if token := firstHeader(headers, "x-hypercube-idp-access-token"); token != "" {
		claims, err := decodeJWT(token)
		if err != nil || tokenExpired(claims, skew) {
			expired = append(expired, "x-hypercube-idp-access-token")
		}
	}
	if len(expired) > 0 {
		return fmt.Errorf("token(s) expirado(s) ou inválido(s): %s", strings.Join(expired, ", "))
	}
	return nil
}

func firstHeader(headers map[string]string, target string) string {
	for key, value := range headers {
		if strings.EqualFold(key, target) {
			return value
		}
	}
	return ""
}

func decodeJWT(token string) (jwtClaims, error) {
	parts := strings.Split(token, ".")
	if len(parts) != 3 {
		return jwtClaims{}, errors.New("jwt inválido")
	}
	payload := parts[1]
	if pad := len(payload) % 4; pad != 0 {
		payload += strings.Repeat("=", 4-pad)
	}
	raw, err := base64.URLEncoding.DecodeString(payload)
	if err != nil {
		return jwtClaims{}, err
	}
	var claims jwtClaims
	if err := json.Unmarshal(raw, &claims); err != nil {
		return jwtClaims{}, err
	}
	return claims, nil
}

func tokenExpired(claims jwtClaims, skew time.Duration) bool {
	if claims.EXP == 0 {
		return true
	}
	return time.Unix(claims.EXP, 0).Before(time.Now().Add(skew))
}

func NewAPIClient(cfg Config, headers map[string]string, logger *Logger) *APIClient {
	transport := &http.Transport{
		Proxy: http.ProxyFromEnvironment,
		DialContext: (&net.Dialer{
			Timeout:   8 * time.Second,
			KeepAlive: 30 * time.Second,
		}).DialContext,
		ForceAttemptHTTP2:     true,
		MaxIdleConns:          32,
		MaxIdleConnsPerHost:   16,
		IdleConnTimeout:       90 * time.Second,
		TLSHandshakeTimeout:   8 * time.Second,
		ExpectContinueTimeout: 1 * time.Second,
	}

	return &APIClient{
		baseURL: docsBase,
		client: &http.Client{
			Timeout:   cfg.RequestTimeout,
			Transport: transport,
		},
		headers: headers,
		limiter: NewRateLimiter(cfg.RequestsPerSec),
		logger:  logger,
	}
}

func (c *APIClient) Close() {
	if c.limiter != nil {
		c.limiter.Close()
	}
}

func NewRateLimiter(rps float64) *RateLimiter {
	if rps <= 0 {
		ch := make(chan time.Time)
		close(ch)
		return &RateLimiter{ch: ch, stop: func() {}}
	}
	interval := time.Duration(float64(time.Second) / rps)
	if interval < 50*time.Millisecond {
		interval = 50 * time.Millisecond
	}
	ticker := time.NewTicker(interval)
	return &RateLimiter{
		ch:   ticker.C,
		stop: ticker.Stop,
	}
}

func (r *RateLimiter) Wait(ctx context.Context) error {
	select {
	case <-ctx.Done():
		return ctx.Err()
	case <-r.ch:
		return nil
	}
}

func (r *RateLimiter) Close() {
	if r != nil && r.stop != nil {
		r.stop()
	}
}

func (c *APIClient) get(ctx context.Context, path string) (*http.Response, []byte, error) {
	if err := c.limiter.Wait(ctx); err != nil {
		return nil, nil, err
	}
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, c.baseURL+path, nil)
	if err != nil {
		return nil, nil, err
	}
	for key, value := range c.headers {
		req.Header.Set(key, value)
	}
	req.Header.Del("If-None-Match")
	req.Header.Del("If-Modified-Since")

	resp, err := c.client.Do(req)
	if err != nil {
		return nil, nil, err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, nil, err
	}
	return resp, body, nil
}

func fetchWorkflowList(ctx context.Context, client *APIClient, retries int, logger *Logger) ([]map[string]any, error) {
	const path = "/workflows/processes?processState=2&requestStateType=3"
	logger.Log(LogContext{}, "Requisitando lista de workflows em: %s%s", docsBase, path)

	for attempt := 1; attempt <= retries; attempt++ {
		resp, body, err := client.get(ctx, path)
		if err != nil {
			logger.Log(LogContext{}, "Erro ao requisitar workflows (tentativa %d/%d): %s", attempt, retries, formatErr(err))
			if attempt < retries {
				time.Sleep(backoff(attempt))
				continue
			}
			return nil, err
		}

		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			logger.Log(LogContext{}, "HTTP %d ao requisitar workflows (tentativa %d/%d).", resp.StatusCode, attempt, retries)
			if isTransientStatus(resp.StatusCode) && attempt < retries {
				time.Sleep(backoff(attempt))
				continue
			}
			return nil, fmt.Errorf("lista de workflows retornou status %d", resp.StatusCode)
		}

		items, err := extractWorkflowItems(body)
		if err != nil {
			logger.Log(LogContext{}, "Erro ao parsear JSON da lista de workflows: %s", formatErr(err))
			return nil, err
		}

		logger.Log(LogContext{}, "Total de workflows retornados pela API: %d", len(items))
		return items, nil
	}
	return nil, fmt.Errorf("não foi possível listar workflows")
}

func extractWorkflowItems(body []byte) ([]map[string]any, error) {
	var data any
	if err := json.Unmarshal(body, &data); err != nil {
		return nil, err
	}
	return anyToWorkflowMaps(data), nil
}

func anyToWorkflowMaps(data any) []map[string]any {
	if data == nil {
		return nil
	}
	if arr, ok := data.([]any); ok {
		result := make([]map[string]any, 0, len(arr))
		for _, item := range arr {
			if m, ok := item.(map[string]any); ok {
				result = append(result, m)
			}
		}
		return result
	}
	if m, ok := data.(map[string]any); ok {
		for _, key := range []string{"items", "data", "results", "Processes", "workflows", "value", "requests"} {
			if arr := anyToWorkflowMaps(m[key]); len(arr) > 0 {
				return arr
			}
		}
	}
	return nil
}

func buildWorkflowGroups(items []map[string]any, targetPeriod string, logger *Logger) []WorkflowGroup {
	grouped := map[string]*WorkflowGroup{}
	for idx, item := range items {
		stub := makeWorkflowStub(item)
		ctx := LogContext{Group: fmt.Sprintf("wf#%d", idx+1), Workflow: stub.ID, Empresa: stub.Company}
		logger.Log(ctx, "Processando workflow id=%s, state=%d", stub.ID, stub.State)

		if stub.State != 2 {
			logger.Log(ctx, "Ignorando workflow (state != 2).")
			continue
		}
		if stub.Category != "SPED" || stub.DocType != "EFD_CONTRIBUICOES" {
			logger.Log(ctx, "Workflow não é SPED EFD_CONTRIBUICOES; ignorando.")
			continue
		}
		if stub.Period != targetPeriod {
			logger.Log(ctx, "Período %s diferente do alvo %s; ignorando.", stub.Period, targetPeriod)
			continue
		}
		if stub.CNPJ == "" {
			logger.Log(ctx, "Workflow sem CNPJ; ignorando.")
			continue
		}

		key := strings.Join([]string{stub.CNPJ, stub.Period, stub.DocType}, "|")
		group := grouped[key]
		if group == nil {
			group = &WorkflowGroup{
				Key:      key,
				Company:  stub.Company,
				CNPJ:     stub.CNPJ,
				Period:   stub.Period,
				Category: stub.Category,
				DocType:  stub.DocType,
			}
			grouped[key] = group
		}
		group.Workflows = append(group.Workflows, stub)
		logger.Log(ctx, "Workflow válido para %s (CNPJ=%s, período=%s).", stub.Company, stub.CNPJ, stub.Period)
	}

	groups := make([]WorkflowGroup, 0, len(grouped))
	for _, group := range grouped {
		sort.Slice(group.Workflows, func(i, j int) bool {
			a := workflowSortTime(group.Workflows[i])
			b := workflowSortTime(group.Workflows[j])
			if a.Equal(b) {
				return group.Workflows[i].ID < group.Workflows[j].ID
			}
			return a.After(b)
		})
		groups = append(groups, *group)
	}
	sort.Slice(groups, func(i, j int) bool {
		return groups[i].Company < groups[j].Company
	})
	return groups
}

func workflowSortTime(stub WorkflowStub) time.Time {
	if !stub.EndDate.IsZero() {
		return stub.EndDate
	}
	return stub.CreationDate
}

func makeWorkflowStub(raw map[string]any) WorkflowStub {
	req := asMap(raw["Request"])
	documents := asSlice(req["Documents"])
	doc0 := asMap(firstSliceItem(documents))
	meta := asMap(asMap(doc0["metadata"])["fileMetadata"])
	clientInfo := asMap(firstNonNil(req["Client"], raw["Client"], doc0["ownerClient"]))

	return WorkflowStub{
		ID:           firstString(raw["id"], raw["sanid"]),
		State:        asInt(raw["state"]),
		Category:     asString(doc0["category"]),
		DocType:      asString(doc0["type"]),
		Company:      firstString(clientInfo["fullName"], strings.TrimSpace(asString(clientInfo["firstName"])+" "+asString(clientInfo["lastName"]))),
		CNPJ:         digitsOnly(firstString(meta["cnpj"], meta["taxpayerNumber"], clientInfo["cpfOrCnpj"])),
		Period:       asString(meta["period"]),
		EndDate:      parseTime(firstString(req["endDate"], raw["endDate"])),
		CreationDate: parseTime(firstString(req["creationDate"], raw["creationDate"])),
		Raw:          raw,
	}
}

func processGroups(ctx context.Context, cfg Config, client *APIClient, logger *Logger, failures *FailureCollector, groups []WorkflowGroup) ProcessSummary {
	jobs := make(chan WorkflowGroup)
	var wg sync.WaitGroup
	var summaryMu sync.Mutex
	summary := ProcessSummary{}

	for workerID := 1; workerID <= cfg.Workers; workerID++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()
			for group := range jobs {
				if err := processGroup(ctx, cfg, client, logger, failures, group, workerID); err != nil {
					summaryMu.Lock()
					summary.Failures++
					summaryMu.Unlock()
				} else {
					summaryMu.Lock()
					summary.Success++
					summaryMu.Unlock()
				}
			}
		}(workerID)
	}

	for _, group := range groups {
		jobs <- group
	}
	close(jobs)
	wg.Wait()
	return summary
}

func processGroup(ctx context.Context, cfg Config, client *APIClient, logger *Logger, failures *FailureCollector, group WorkflowGroup, workerID int) error {
	ctxLog := LogContext{
		Group:    group.CNPJ + "/" + group.Period,
		Workflow: fmt.Sprintf("worker-%d", workerID),
		Empresa:  group.Company,
	}

	if len(group.Workflows) > 1 {
		logger.Log(ctxLog, "Grupo deduplicado com %d workflows candidatos.", len(group.Workflows))
	}

	outputPath := buildOutputPath(cfg.OutputDir, group.Company, group.CNPJ, group.Period)
	if cfg.SkipExisting {
		if info, err := os.Stat(outputPath); err == nil && info.Size() > 0 {
			if info.Size() < 2500 {
				logger.Log(ctxLog, "Arquivo existente parece suspeito (%d bytes); reprocessando: %s", info.Size(), outputPath)
			} else {
				logger.Log(ctxLog, "Arquivo já existe; pulando download: %s", outputPath)
				return nil
			}
		}
	}

	var lastErr error
	var pendingFailure *FailureEntry
	for _, wf := range group.Workflows {
		wfCtx := ctxLog
		wfCtx.Workflow = wf.ID

		details, err := fetchFullDetails(ctx, client, logger, wfCtx, wf.ID, cfg.DetailRetries)
		if err != nil {
			lastErr = err
			entry := FailureEntry{
				GroupKey:   group.Key,
				Empresa:    group.Company,
				CNPJ:       group.CNPJ,
				Periodo:    group.Period,
				WorkflowID: wf.ID,
				Reason:     err.Error(),
			}
			pendingFailure = &entry
			continue
		}

		doc, reason, ok := selectBestReceiptDocument(details, group.CNPJ, group.Period)
		if !ok {
			if reason == "" {
				reason = "nenhum recibo de entrega disponível"
			}
			logger.Log(wfCtx, "%s", reason)
			lastErr = errors.New(reason)
			entry := FailureEntry{
				GroupKey:   group.Key,
				Empresa:    group.Company,
				CNPJ:       group.CNPJ,
				Periodo:    group.Period,
				WorkflowID: wf.ID,
				Reason:     reason,
			}
			pendingFailure = &entry
			continue
		}

		logger.Log(wfCtx, "Documento escolhido: id=%s, score=%d, kind=%s, name=%s", doc.ID, doc.Score, doc.MatchKind, doc.Name)
		result, err := downloadDocument(ctx, cfg, client, logger, wfCtx, doc, outputPath)
		if err != nil {
			lastErr = err
			entry := FailureEntry{
				GroupKey:   group.Key,
				Empresa:    group.Company,
				CNPJ:       group.CNPJ,
				Periodo:    group.Period,
				WorkflowID: wf.ID,
				DocumentID: doc.ID,
				Document:   doc.Name,
				OutputPath: outputPath,
				Reason:     err.Error(),
			}
			pendingFailure = &entry
			continue
		}

		logger.Log(wfCtx, "PDF salvo com sucesso: %s", result.Path)
		return nil
	}

	if lastErr == nil {
		lastErr = errors.New("falha sem detalhe")
	}
	if pendingFailure != nil {
		failures.Add(*pendingFailure)
	}
	logger.Log(ctxLog, "Falha final do grupo: %s", lastErr.Error())
	return lastErr
}

func fetchFullDetails(ctx context.Context, client *APIClient, logger *Logger, logCtx LogContext, workflowID string, retries int) (map[string]any, error) {
	path := fmt.Sprintf("/workflows/processes/%s?fullDetails=true", workflowID)
	for attempt := 1; attempt <= retries; attempt++ {
		resp, body, err := client.get(ctx, path)
		if err != nil {
			logger.Log(logCtx, "Erro ao chamar fullDetails (%s) tentativa %d/%d: %s", workflowID, attempt, retries, formatErr(err))
			if attempt < retries {
				time.Sleep(backoff(attempt))
				continue
			}
			return nil, err
		}

		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			logger.Log(logCtx, "fullDetails retornou HTTP %d para %s.", resp.StatusCode, workflowID)
			if isTransientStatus(resp.StatusCode) && attempt < retries {
				time.Sleep(backoff(attempt))
				continue
			}
			return nil, fmt.Errorf("fullDetails HTTP %d", resp.StatusCode)
		}

		var data map[string]any
		if err := json.Unmarshal(body, &data); err != nil {
			logger.Log(logCtx, "Erro ao parsear JSON de fullDetails para %s: %s", workflowID, formatErr(err))
			if attempt < retries {
				time.Sleep(backoff(attempt))
				continue
			}
			return nil, err
		}
		return data, nil
	}
	return nil, fmt.Errorf("fullDetails falhou para %s", workflowID)
}

func selectBestReceiptDocument(details map[string]any, cnpj, period string) (DocumentCandidate, string, bool) {
	var candidates []DocumentCandidate
	seen := map[string]struct{}{}
	collectDocumentCandidates(details, cnpj, period, seen, &candidates)
	if len(candidates) == 0 {
		return DocumentCandidate{}, "nenhum PDF candidato encontrado no workflow", false
	}

	sort.Slice(candidates, func(i, j int) bool {
		if candidates[i].Score == candidates[j].Score {
			return candidates[i].Name < candidates[j].Name
		}
		return candidates[i].Score > candidates[j].Score
	})

	for _, candidate := range candidates {
		if candidate.MatchKind == "receipt" {
			return candidate, "", true
		}
	}

	for _, candidate := range candidates {
		if candidate.MatchKind == "comprovante" {
			return DocumentCandidate{}, fmt.Sprintf("somente comprovante disponível no workflow: %s", candidate.Name), false
		}
	}

	names := make([]string, 0, len(candidates))
	for _, candidate := range candidates {
		names = append(names, candidate.Name)
	}
	return DocumentCandidate{}, "nenhum Recibo de Entrega disponível; candidatos encontrados: " + strings.Join(names, "; "), false
}

func collectDocumentCandidates(node any, cnpj, period string, seen map[string]struct{}, out *[]DocumentCandidate) {
	switch value := node.(type) {
	case map[string]any:
		if docMap, ok := value["Document"].(map[string]any); ok {
			if candidate, ok := makeDocumentCandidate(docMap, cnpj, period); ok {
				key := candidate.ID + "|" + candidate.Name
				if _, exists := seen[key]; !exists {
					seen[key] = struct{}{}
					*out = append(*out, candidate)
				}
			}
		}
		if candidate, ok := makeDocumentCandidate(value, cnpj, period); ok {
			key := candidate.ID + "|" + candidate.Name
			if _, exists := seen[key]; !exists {
				seen[key] = struct{}{}
				*out = append(*out, candidate)
			}
		}
		for _, child := range value {
			collectDocumentCandidates(child, cnpj, period, seen, out)
		}
	case []any:
		for _, child := range value {
			collectDocumentCandidates(child, cnpj, period, seen, out)
		}
	}
}

func makeDocumentCandidate(doc map[string]any, cnpj, period string) (DocumentCandidate, bool) {
	id := asString(doc["id"])
	name := asString(doc["name"])
	if id == "" || name == "" {
		return DocumentCandidate{}, false
	}

	ext := strings.ToLower(asString(doc["extension"]))
	mime := strings.ToLower(asString(doc["mimeType"]))
	isPDF := asBool(doc["isPdf"]) || ext == "pdf" || strings.Contains(mime, "pdf")
	if !isPDF {
		return DocumentCandidate{}, false
	}

	score, matchKind, ok := scoreDocument(name, cnpj, period)
	if !ok {
		return DocumentCandidate{}, false
	}

	return DocumentCandidate{
		ID:        id,
		Name:      name,
		Extension: ext,
		MimeType:  mime,
		Score:     score,
		MatchKind: matchKind,
	}, true
}

func scoreDocument(name, cnpj, period string) (int, string, bool) {
	norm := normalizeForMatch(name)
	cleanCNPJ := digitsOnly(cnpj)
	score := 0
	matchKind := ""

	if strings.Contains(norm, "pendencia") || strings.Contains(norm, "validacao") || strings.Contains(norm, "erro") {
		return 0, "", false
	}

	hasRecibo := strings.Contains(norm, "recibo")
	hasEntrega := strings.Contains(norm, "entrega")
	hasComprovante := strings.Contains(norm, "comprovante")

	if hasRecibo && hasEntrega {
		score += 200
		matchKind = "receipt"
	} else if hasComprovante {
		score += 100
		matchKind = "comprovante"
	} else {
		return 0, "", false
	}
	if strings.Contains(norm, "entrega") {
		score += 35
	}
	if strings.Contains(norm, "transmissao") {
		score += 30
	}
	if strings.Contains(norm, "efd contribuicoes") || strings.Contains(norm, "efd_contribuicoes") {
		score += 15
	}
	if cleanCNPJ != "" && strings.Contains(norm, cleanCNPJ) {
		score += 60
	}
	for _, token := range periodTokens(period) {
		if token != "" && strings.Contains(norm, token) {
			score += 40
			break
		}
	}

	if cleanCNPJ != "" && !strings.Contains(norm, cleanCNPJ) {
		score -= 20
	}

	switch matchKind {
	case "receipt":
		return score, matchKind, true
	case "comprovante":
		return score, matchKind, true
	default:
		return 0, "", false
	}
}

func normalizeForMatch(s string) string {
	s = strings.ToLower(s)
	replacer := strings.NewReplacer(
		"á", "a", "à", "a", "â", "a", "ã", "a",
		"é", "e", "ê", "e",
		"í", "i",
		"ó", "o", "ô", "o", "õ", "o",
		"ú", "u",
		"ç", "c",
		"_", " ", "-", " ", "/", " ", ".", " ",
	)
	s = replacer.Replace(s)
	fields := strings.Fields(s)
	return strings.Join(fields, " ")
}

func periodTokens(period string) []string {
	if len(period) != 6 {
		return []string{period}
	}
	month := period[:2]
	year := period[2:]
	return []string{
		period,
		month + year,
		month + "-" + year,
		month + "/" + year,
		month + " " + year,
	}
}

func downloadDocument(ctx context.Context, cfg Config, client *APIClient, logger *Logger, logCtx LogContext, doc DocumentCandidate, outputPath string) (DownloadResult, error) {
	path := fmt.Sprintf("/documents/%s/download", doc.ID)
	emptyResponseCount := 0
	lastStatus := 0

	for attempt := 1; attempt <= cfg.DownloadRetries; attempt++ {
		logger.Log(logCtx, "download: GET %s%s (tentativa %d/%d)", docsBase, path, attempt, cfg.DownloadRetries)

		resp, body, err := client.get(ctx, path)
		if err != nil {
			logger.Log(logCtx, "download: erro de rede: %s", formatErr(err))
			if attempt < cfg.DownloadRetries {
				delay := backoff(attempt)
				logger.Log(logCtx, "download: aguardando %s antes de nova tentativa...", delay)
				time.Sleep(delay)
				continue
			}
			return DownloadResult{}, err
		}

		ct := strings.ToLower(resp.Header.Get("Content-Type"))
		lastStatus = resp.StatusCode
		logger.Log(logCtx, "download: status=%d, content-type=%s, bytes=%d", resp.StatusCode, ct, len(body))

		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			logger.Log(logCtx, "download: corpo (início): %s", previewBytes(body, 240))
			if isTransientStatus(resp.StatusCode) && attempt < cfg.DownloadRetries {
				delay := backoff(attempt)
				logger.Log(logCtx, "download: erro transitório (%d). Aguardando %s...", resp.StatusCode, delay)
				time.Sleep(delay)
				continue
			}
			return DownloadResult{}, fmt.Errorf("download HTTP %d para documento %s", resp.StatusCode, doc.ID)
		}

		if !looksLikePDF(body, ct) {
			if resp.StatusCode == http.StatusNoContent && len(body) == 0 {
				emptyResponseCount++
				if attempt == cfg.DownloadRetries && lastStatus == http.StatusNoContent {
					return DownloadResult{}, fmt.Errorf("documento %s indisponivel para download na API (HTTP 204 repetido)", doc.ID)
				}
			}
			logger.Log(logCtx, "download: conteúdo inválido para PDF. preview=%q", previewBytes(body, 160))
			if attempt < cfg.DownloadRetries {
				delay := backoff(attempt)
				logger.Log(logCtx, "download: aguardando %s antes de nova tentativa...", delay)
				time.Sleep(delay)
				continue
			}
			return DownloadResult{}, fmt.Errorf("resposta sem PDF válido para documento %s", doc.ID)
		}

		if err := writeFileAtomically(outputPath, body); err != nil {
			if attempt < cfg.DownloadRetries {
				delay := backoff(attempt)
				logger.Log(logCtx, "download: erro ao salvar arquivo: %s. Aguardando %s...", formatErr(err), delay)
				time.Sleep(delay)
				continue
			}
			return DownloadResult{}, err
		}

		return DownloadResult{
			Path:         outputPath,
			DocumentID:   doc.ID,
			DocumentName: doc.Name,
		}, nil
	}

	if emptyResponseCount > 0 && lastStatus == http.StatusNoContent {
		return DownloadResult{}, fmt.Errorf("documento %s indisponivel para download na API (HTTP 204 repetido)", doc.ID)
	}
	return DownloadResult{}, fmt.Errorf("download falhou para documento %s", doc.ID)
}

func looksLikePDF(body []byte, contentType string) bool {
	if len(body) == 0 {
		return false
	}
	if bytes.HasPrefix(body, []byte("%PDF-")) {
		return true
	}
	if strings.Contains(contentType, "application/pdf") && len(body) > 512 {
		return true
	}
	if strings.Contains(contentType, "application/octet-stream") && len(body) > 1024 {
		lower := strings.ToLower(string(body[:min(len(body), 256)]))
		return !strings.Contains(lower, "<html") && !strings.Contains(lower, "sem resposta do servidor")
	}
	return false
}

func writeFileAtomically(path string, body []byte) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	tmp := path + ".tmp"
	if err := os.WriteFile(tmp, body, 0o644); err != nil {
		return err
	}
	return os.Rename(tmp, path)
}

func buildOutputPath(outDir, company, cnpj, period string) string {
	name := sanitizeFilename(fmt.Sprintf("%s - %s - %s EFD CONTRIBUICOES.pdf", company, cnpj, period))
	return filepath.Join(outDir, name)
}

func sanitizeFilename(name string) string {
	replacer := strings.NewReplacer("<", "_", ">", "_", ":", "_", "\"", "_", "/", "_", "\\", "_", "|", "_", "?", "_", "*", "_")
	name = replacer.Replace(strings.TrimSpace(name))
	name = strings.Join(strings.Fields(name), " ")
	if name == "" {
		return "arquivo.pdf"
	}
	return name
}

func previewBytes(body []byte, max int) string {
	if len(body) > max {
		body = body[:max]
	}
	return strings.TrimSpace(string(body))
}

func saveFailures(baseDir string, logger *Logger, failures *FailureCollector) {
	entries := failures.List()
	if len(entries) == 0 {
		logger.Log(LogContext{}, "Nenhum download falhou. Nenhum relatório de falhas gerado.")
		return
	}

	timestamp := time.Now().Format("20060102-150405")
	jsonPath := filepath.Join(baseDir, "falhas_download_go_"+timestamp+".json")
	csvPath := filepath.Join(baseDir, "falhas_download_go_"+timestamp+".csv")

	if raw, err := json.MarshalIndent(entries, "", "  "); err == nil {
		if err := os.WriteFile(jsonPath, raw, 0o644); err == nil {
			logger.Log(LogContext{}, "Relatório de falhas (JSON) salvo em: %s", jsonPath)
		} else {
			logger.Log(LogContext{}, "Erro ao salvar relatório JSON: %s", formatErr(err))
		}
	} else {
		logger.Log(LogContext{}, "Erro ao serializar relatório JSON: %s", formatErr(err))
	}

	file, err := os.Create(csvPath)
	if err != nil {
		logger.Log(LogContext{}, "Erro ao criar CSV de falhas: %s", formatErr(err))
		return
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	header := []string{"group_key", "empresa", "cnpj", "periodo", "workflow_id", "document_id", "document_name", "output_path", "reason"}
	if err := writer.Write(header); err != nil {
		logger.Log(LogContext{}, "Erro ao escrever cabeçalho CSV: %s", formatErr(err))
		return
	}
	for _, entry := range entries {
		row := []string{
			entry.GroupKey,
			entry.Empresa,
			entry.CNPJ,
			entry.Periodo,
			entry.WorkflowID,
			entry.DocumentID,
			entry.Document,
			entry.OutputPath,
			entry.Reason,
		}
		if err := writer.Write(row); err != nil {
			logger.Log(LogContext{}, "Erro ao escrever linha CSV: %s", formatErr(err))
			return
		}
	}
	logger.Log(LogContext{}, "Relatório de falhas (CSV) salvo em: %s", csvPath)
}

func (f *FailureCollector) Add(entry FailureEntry) {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.entries = append(f.entries, entry)
}

func (f *FailureCollector) List() []FailureEntry {
	f.mu.Lock()
	defer f.mu.Unlock()
	out := make([]FailureEntry, len(f.entries))
	copy(out, f.entries)
	return out
}

func backoff(attempt int) time.Duration {
	seconds := math.Pow(2, float64(attempt-1))
	if seconds > 8 {
		seconds = 8
	}
	return time.Duration(seconds * float64(time.Second))
}

func isTransientStatus(status int) bool {
	_, ok := transientStatusCodes[status]
	return ok
}

func formatErr(err error) string {
	if err == nil {
		return ""
	}
	return fmt.Sprintf("%T: %v", err, err)
}

func first(values ...any) any {
	for _, value := range values {
		if value != nil {
			return value
		}
	}
	return nil
}

func firstNonNil(values ...any) any {
	return first(values...)
}

func firstSliceItem(values []any) any {
	if len(values) == 0 {
		return nil
	}
	return values[0]
}

func firstString(values ...any) string {
	for _, value := range values {
		if s := asString(value); s != "" {
			return s
		}
	}
	return ""
}

func asMap(value any) map[string]any {
	if value == nil {
		return map[string]any{}
	}
	if m, ok := value.(map[string]any); ok {
		return m
	}
	return map[string]any{}
}

func asSlice(value any) []any {
	if value == nil {
		return nil
	}
	if arr, ok := value.([]any); ok {
		return arr
	}
	return nil
}

func asString(value any) string {
	switch v := value.(type) {
	case nil:
		return ""
	case string:
		return strings.TrimSpace(v)
	case float64:
		if v == math.Trunc(v) {
			return strconv.FormatInt(int64(v), 10)
		}
		return strconv.FormatFloat(v, 'f', -1, 64)
	case json.Number:
		return v.String()
	case bool:
		return strconv.FormatBool(v)
	default:
		return strings.TrimSpace(fmt.Sprint(v))
	}
}

func asInt(value any) int {
	switch v := value.(type) {
	case nil:
		return 0
	case int:
		return v
	case int64:
		return int(v)
	case float64:
		return int(v)
	case json.Number:
		n, _ := v.Int64()
		return int(n)
	case string:
		n, _ := strconv.Atoi(strings.TrimSpace(v))
		return n
	default:
		return 0
	}
}

func asBool(value any) bool {
	switch v := value.(type) {
	case bool:
		return v
	case string:
		b, _ := strconv.ParseBool(strings.TrimSpace(v))
		return b
	default:
		return false
	}
}

func parseTime(value string) time.Time {
	if value == "" {
		return time.Time{}
	}
	formats := []string{
		time.RFC3339,
		"2006-01-02T15:04:05.000Z",
		"2006-01-02T15:04:05Z",
	}
	for _, format := range formats {
		if t, err := time.Parse(format, value); err == nil {
			return t
		}
	}
	return time.Time{}
}

func digitsOnly(s string) string {
	var b strings.Builder
	for _, r := range s {
		if r >= '0' && r <= '9' {
			b.WriteRune(r)
		}
	}
	return b.String()
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
