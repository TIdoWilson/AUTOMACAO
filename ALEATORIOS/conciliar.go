package main

import (
	"bufio"
	"fmt"
	"math/bits"
	"os"
	"sort"
	"strconv"
	"strings"
)

type Item struct {
	Index int
	Line  int
	Value int64
	Used  bool
}

type PairCombo struct {
	I, J int
	Sum  int64
}

type TripleCombo struct {
	I, J, K int
	Sum     int64
}

type Candidate struct {
	AIdx    []int
	BIdx    []int
	Sum     int64
	Pattern string
	Ok      bool
}

func parseBRLToCents(s string) (int64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("valor vazio")
	}

	neg := false
	if strings.HasPrefix(s, "-") {
		neg = true
		s = s[1:]
	}

	s = strings.ReplaceAll(s, ".", "")
	parts := strings.Split(s, ",")

	intPart := "0"
	fracPart := "00"

	if len(parts) >= 1 && parts[0] != "" {
		intPart = parts[0]
	}
	if len(parts) >= 2 {
		fracPart = parts[1]
	}
	if len(fracPart) == 1 {
		fracPart += "0"
	}
	if len(fracPart) > 2 {
		fracPart = fracPart[:2]
	}

	i, err := strconv.ParseInt(intPart, 10, 64)
	if err != nil {
		return 0, err
	}
	f, err := strconv.ParseInt(fracPart, 10, 64)
	if err != nil {
		return 0, err
	}

	v := i*100 + f
	if neg {
		v = -v
	}
	return v, nil
}

func formatBRL(cents int64) string {
	neg := cents < 0
	if neg {
		cents = -cents
	}

	inteiro := cents / 100
	frac := cents % 100

	s := strconv.FormatInt(inteiro, 10)
	var out []byte
	count := 0

	for i := len(s) - 1; i >= 0; i-- {
		out = append(out, s[i])
		count++
		if count == 3 && i > 0 {
			out = append(out, '.')
			count = 0
		}
	}

	for i, j := 0, len(out)-1; i < j; i, j = i+1, j-1 {
		out[i], out[j] = out[j], out[i]
	}

	res := string(out) + fmt.Sprintf(",%02d", frac)
	if neg {
		return "-" + res
	}
	return res
}

func loadItems(filename string) ([]Item, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	var items []Item
	scanner := bufio.NewScanner(f)
	line := 0

	for scanner.Scan() {
		line++
		txt := strings.TrimSpace(scanner.Text())
		if txt == "" {
			continue
		}

		v, err := parseBRLToCents(txt)
		if err != nil {
			return nil, fmt.Errorf("%s linha %d: %w", filename, line, err)
		}

		items = append(items, Item{
			Index: len(items),
			Line:  line,
			Value: v,
			Used:  false,
		})
	}

	if err := scanner.Err(); err != nil {
		return nil, err
	}

	return items, nil
}

func buildSortedOrder(items []Item) []int {
	order := make([]int, len(items))
	for i := range items {
		order[i] = i
	}

	sort.Slice(order, func(i, j int) bool {
		vi := items[order[i]].Value
		vj := items[order[j]].Value
		if vi != vj {
			return vi > vj
		}
		return items[order[i]].Line < items[order[j]].Line
	})

	return order
}

func buildPairsSorted(items []Item) []PairCombo {
	n := len(items)
	out := make([]PairCombo, 0, n*(n-1)/2)

	for i := 0; i < n; i++ {
		for j := i + 1; j < n; j++ {
			out = append(out, PairCombo{
				I:   i,
				J:   j,
				Sum: items[i].Value + items[j].Value,
			})
		}
	}

	sort.Slice(out, func(i, j int) bool {
		if out[i].Sum != out[j].Sum {
			return out[i].Sum > out[j].Sum
		}
		if out[i].I != out[j].I {
			return out[i].I < out[j].I
		}
		return out[i].J < out[j].J
	})

	return out
}

func buildTriplesSorted(items []Item) []TripleCombo {
	n := len(items)
	out := make([]TripleCombo, 0)

	for i := 0; i < n; i++ {
		for j := i + 1; j < n; j++ {
			for k := j + 1; k < n; k++ {
				out = append(out, TripleCombo{
					I:   i,
					J:   j,
					K:   k,
					Sum: items[i].Value + items[j].Value + items[k].Value,
				})
			}
		}
	}

	sort.Slice(out, func(i, j int) bool {
		if out[i].Sum != out[j].Sum {
			return out[i].Sum > out[j].Sum
		}
		if out[i].I != out[j].I {
			return out[i].I < out[j].I
		}
		if out[i].J != out[j].J {
			return out[i].J < out[j].J
		}
		return out[i].K < out[j].K
	})

	return out
}

func packPair(i, j int) uint64 {
	return (uint64(uint32(i)) << 32) | uint64(uint32(j))
}

func unpackPair(v uint64) (int, int) {
	return int(uint32(v >> 32)), int(uint32(v))
}

func buildPairBuckets(items []Item, maxUsefulSum int64) map[int64][]uint64 {
	n := len(items)
	buckets := make(map[int64][]uint64)

	for i := 0; i < n; i++ {
		vi := items[i].Value
		for j := i + 1; j < n; j++ {
			sum := vi + items[j].Value
			if maxUsefulSum > 0 && sum > maxUsefulSum {
				continue
			}
			buckets[sum] = append(buckets[sum], packPair(i, j))
		}
	}

	return buckets
}

func buildSingleMapUnused(items []Item, order []int) map[int64]int {
	m := make(map[int64]int)
	for _, idx := range order {
		if items[idx].Used {
			continue
		}
		v := items[idx].Value
		if _, ok := m[v]; !ok {
			m[v] = idx
		}
	}
	return m
}

func firstValidPair(bucket []uint64, items []Item, forbid ...int) (int, int, bool) {
	if len(bucket) == 0 {
		return 0, 0, false
	}

	var forbidMap map[int]struct{}
	if len(forbid) > 0 {
		forbidMap = make(map[int]struct{}, len(forbid))
		for _, f := range forbid {
			forbidMap[f] = struct{}{}
		}
	}

	for _, packed := range bucket {
		i, j := unpackPair(packed)

		if items[i].Used || items[j].Used {
			continue
		}
		if forbidMap != nil {
			if _, ok := forbidMap[i]; ok {
				continue
			}
			if _, ok := forbidMap[j]; ok {
				continue
			}
		}
		return i, j, true
	}

	return 0, 0, false
}

func validPair(p PairCombo, items []Item) bool {
	return !items[p.I].Used && !items[p.J].Used
}

func validTriple(t TripleCombo, items []Item) bool {
	return !items[t.I].Used && !items[t.J].Used && !items[t.K].Used
}

func better(a, b Candidate) bool {
	if !a.Ok {
		return false
	}
	if !b.Ok {
		return true
	}
	if a.Sum != b.Sum {
		return a.Sum > b.Sum
	}
	if len(a.AIdx)+len(a.BIdx) != len(b.AIdx)+len(b.BIdx) {
		return len(a.AIdx)+len(a.BIdx) < len(b.AIdx)+len(b.BIdx)
	}
	return a.Pattern < b.Pattern
}

func find1x1(list1, list2 []Item, order1, order2 []int) Candidate {
	single2 := buildSingleMapUnused(list2, order2)
	for _, i := range order1 {
		if list1[i].Used {
			continue
		}
		if j, ok := single2[list1[i].Value]; ok {
			return Candidate{
				AIdx:    []int{i},
				BIdx:    []int{j},
				Sum:     list1[i].Value,
				Pattern: "1x1",
				Ok:      true,
			}
		}
	}
	return Candidate{}
}

func find1x2(list1, list2 []Item, order1 []int, pairs2 []PairCombo) Candidate {
	single1 := buildSingleMapUnused(list1, order1)
	for _, p := range pairs2 {
		if !validPair(p, list2) {
			continue
		}
		if i, ok := single1[p.Sum]; ok {
			return Candidate{
				AIdx:    []int{i},
				BIdx:    []int{p.I, p.J},
				Sum:     p.Sum,
				Pattern: "1x2",
				Ok:      true,
			}
		}
	}
	return Candidate{}
}

func find2x1(list1, list2 []Item, order2 []int, pairBuckets1 map[int64][]uint64) Candidate {
	for _, j := range order2 {
		if list2[j].Used {
			continue
		}
		target := list2[j].Value
		if i1, i2, ok := firstValidPair(pairBuckets1[target], list1); ok {
			return Candidate{
				AIdx:    []int{i1, i2},
				BIdx:    []int{j},
				Sum:     target,
				Pattern: "2x1",
				Ok:      true,
			}
		}
	}
	return Candidate{}
}

func find2x2(list1, list2 []Item, pairs2 []PairCombo, pairBuckets1 map[int64][]uint64) Candidate {
	for _, p := range pairs2 {
		if !validPair(p, list2) {
			continue
		}
		if i1, i2, ok := firstValidPair(pairBuckets1[p.Sum], list1); ok {
			return Candidate{
				AIdx:    []int{i1, i2},
				BIdx:    []int{p.I, p.J},
				Sum:     p.Sum,
				Pattern: "2x2",
				Ok:      true,
			}
		}
	}
	return Candidate{}
}

func find1x3(list1, list2 []Item, order1 []int, triples2 []TripleCombo) Candidate {
	single1 := buildSingleMapUnused(list1, order1)
	for _, t := range triples2 {
		if !validTriple(t, list2) {
			continue
		}
		if i, ok := single1[t.Sum]; ok {
			return Candidate{
				AIdx:    []int{i},
				BIdx:    []int{t.I, t.J, t.K},
				Sum:     t.Sum,
				Pattern: "1x3",
				Ok:      true,
			}
		}
	}
	return Candidate{}
}

func find2x3(list1, list2 []Item, triples2 []TripleCombo, pairBuckets1 map[int64][]uint64) Candidate {
	for _, t := range triples2 {
		if !validTriple(t, list2) {
			continue
		}
		if i1, i2, ok := firstValidPair(pairBuckets1[t.Sum], list1); ok {
			return Candidate{
				AIdx:    []int{i1, i2},
				BIdx:    []int{t.I, t.J, t.K},
				Sum:     t.Sum,
				Pattern: "2x3",
				Ok:      true,
			}
		}
	}
	return Candidate{}
}

func find3x1(list1, list2 []Item, order1, order2 []int, pairBuckets1 map[int64][]uint64) Candidate {
	for _, j := range order2 {
		if list2[j].Used {
			continue
		}
		target := list2[j].Value

		for _, i := range order1 {
			if list1[i].Used {
				continue
			}
			rem := target - list1[i].Value
			if rem <= 0 {
				continue
			}
			if a, b, ok := firstValidPair(pairBuckets1[rem], list1, i); ok {
				return Candidate{
					AIdx:    []int{i, a, b},
					BIdx:    []int{j},
					Sum:     target,
					Pattern: "3x1",
					Ok:      true,
				}
			}
		}
	}
	return Candidate{}
}

func find3x2(list1, list2 []Item, order1 []int, pairs2 []PairCombo, pairBuckets1 map[int64][]uint64) Candidate {
	for _, p := range pairs2 {
		if !validPair(p, list2) {
			continue
		}
		target := p.Sum

		for _, i := range order1 {
			if list1[i].Used {
				continue
			}
			rem := target - list1[i].Value
			if rem <= 0 {
				continue
			}
			if a, b, ok := firstValidPair(pairBuckets1[rem], list1, i); ok {
				return Candidate{
					AIdx:    []int{i, a, b},
					BIdx:    []int{p.I, p.J},
					Sum:     target,
					Pattern: "3x2",
					Ok:      true,
				}
			}
		}
	}
	return Candidate{}
}

func markUsed(items []Item, idxs []int) {
	for _, idx := range idxs {
		items[idx].Used = true
	}
}

func valuesString(items []Item, idxs []int) string {
	type row struct {
		Value int64
		Line  int
	}
	rows := make([]row, 0, len(idxs))
	for _, idx := range idxs {
		rows = append(rows, row{
			Value: items[idx].Value,
			Line:  items[idx].Line,
		})
	}

	sort.Slice(rows, func(i, j int) bool {
		if rows[i].Value != rows[j].Value {
			return rows[i].Value > rows[j].Value
		}
		return rows[i].Line < rows[j].Line
	})

	parts := make([]string, 0, len(rows))
	for _, r := range rows {
		parts = append(parts, fmt.Sprintf("%s (linha %d)", formatBRL(r.Value), r.Line))
	}
	return strings.Join(parts, " + ")
}

func countUnused(items []Item) int {
	c := 0
	for _, it := range items {
		if !it.Used {
			c++
		}
	}
	return c
}

func sumUnused(items []Item) int64 {
	var total int64
	for _, it := range items {
		if !it.Used {
			total += it.Value
		}
	}
	return total
}

func maxUsefulSum(list2 []Item) int64 {
	vals := make([]int64, 0, len(list2))
	for _, it := range list2 {
		vals = append(vals, it.Value)
	}
	sort.Slice(vals, func(i, j int) bool { return vals[i] > vals[j] })

	var total int64
	for i := 0; i < len(vals) && i < 3; i++ {
		total += vals[i]
	}
	return total
}

func printMatches(matches []Candidate, list1, list2 []Item) {
	if len(matches) == 0 {
		fmt.Println("Nenhuma combinação encontrada.")
		return
	}

	var totalConciliado int64
	for i, m := range matches {
		totalConciliado += m.Sum
		fmt.Printf("[%d] Padrão %s | Soma: %s\n", i+1, m.Pattern, formatBRL(m.Sum))
		fmt.Printf("  Lista 1: %s\n", valuesString(list1, m.AIdx))
		fmt.Printf("  Lista 2: %s\n", valuesString(list2, m.BIdx))
		fmt.Println()
	}

	fmt.Printf("Total conciliado: %s\n", formatBRL(totalConciliado))
}

func printUnmatched(title string, items []Item) {
	fmt.Println(title)
	fmt.Println(strings.Repeat("-", len(title)))

	type row struct {
		Line  int
		Value int64
	}
	var rows []row
	for _, it := range items {
		if !it.Used {
			rows = append(rows, row{
				Line:  it.Line,
				Value: it.Value,
			})
		}
	}

	if len(rows) == 0 {
		fmt.Println("Nenhum item sem combinação.")
		fmt.Println()
		return
	}

	sort.Slice(rows, func(i, j int) bool {
		if rows[i].Value != rows[j].Value {
			return rows[i].Value > rows[j].Value
		}
		return rows[i].Line < rows[j].Line
	})

	var total int64
	for _, r := range rows {
		total += r.Value
		fmt.Printf("Linha %d -> %s\n", r.Line, formatBRL(r.Value))
	}

	fmt.Printf("Quantidade sem combinação: %d\n", len(rows))
	fmt.Printf("Total sem combinação: %s\n", formatBRL(total))
	fmt.Println()
}

func askDisplayMode() string {
	reader := bufio.NewReader(os.Stdin)

	for {
		fmt.Println()
		fmt.Println("Escolha o que deseja visualizar:")
		fmt.Println("1 - Combinações encontradas")
		fmt.Println("2 - Itens sem combinação")
		fmt.Println("3 - Ambos")
		fmt.Print("Opção: ")

		input, _ := reader.ReadString('\n')
		input = strings.TrimSpace(input)

		if input == "1" || input == "2" || input == "3" {
			return input
		}

		fmt.Println("Opção inválida. Digite 1, 2 ou 3.")
	}
}

func min64(a, b int64) int64 {
	if a < b {
		return a
	}
	return b
}

func sumAll(items []Item) int64 {
	var total int64
	for _, it := range items {
		total += it.Value
	}
	return total
}

func unusedIndicesSorted(items []Item) []int {
	idxs := make([]int, 0)
	for i := range items {
		if !items[i].Used {
			idxs = append(idxs, i)
		}
	}

	sort.Slice(idxs, func(i, j int) bool {
		vi := items[idxs[i]].Value
		vj := items[idxs[j]].Value
		if vi != vj {
			return vi > vj
		}
		return items[idxs[i]].Line < items[idxs[j]].Line
	})

	return idxs
}

func buildReachableBitset(items []Item, idxs []int, limit int) []uint64 {
	if limit < 0 {
		return nil
	}

	words := make([]uint64, limit/64+1)
	words[0] = 1

	maskLast := func() {
		rem := uint((limit + 1) % 64)
		if rem != 0 {
			words[len(words)-1] &= (uint64(1) << rem) - 1
		}
	}

	for _, idx := range idxs {
		shift := int(items[idx].Value)
		if shift <= 0 || shift > limit {
			continue
		}

		wordShift := shift / 64
		bitShift := uint(shift % 64)

		for i := len(words) - 1; i >= 0; i-- {
			src := i - wordShift
			if src < 0 {
				continue
			}

			var moved uint64
			moved = words[src] << bitShift

			if bitShift > 0 && src-1 >= 0 {
				moved |= words[src-1] >> (64 - bitShift)
			}

			words[i] |= moved
		}

		maskLast()
	}

	return words
}

func largestCommonReachableSum(a, b []uint64, limit int) int {
	if len(a) == 0 || len(b) == 0 {
		return 0
	}

	lastWord := limit / 64
	lastBits := uint((limit % 64) + 1)

	for w := lastWord; w >= 0; w-- {
		inter := a[w] & b[w]

		if w == 0 {
			inter &^= 1
		}

		if w == lastWord && lastBits < 64 {
			inter &= (uint64(1) << lastBits) - 1
		}

		if inter != 0 {
			bit := 63 - bits.LeadingZeros64(inter)
			return w*64 + bit
		}
	}

	return 0
}

func reconstructSubsetForTarget(items []Item, idxs []int, target int) ([]int, bool) {
	if target < 0 {
		return nil, false
	}
	if target == 0 {
		return []int{}, true
	}

	prevSum := make([]int32, target+1)
	picked := make([]int32, target+1)

	for i := 0; i <= target; i++ {
		prevSum[i] = -1
		picked[i] = -1
	}
	prevSum[0] = -2

	reachableMax := 0

	for _, idx := range idxs {
		v := int(items[idx].Value)
		if v <= 0 || v > target {
			continue
		}

		upper := reachableMax + v
		if upper > target {
			upper = target
		}

		for s := upper; s >= v; s-- {
			if prevSum[s] != -1 {
				continue
			}
			if prevSum[s-v] != -1 {
				prevSum[s] = int32(s - v)
				picked[s] = int32(idx)
			}
		}

		if upper > reachableMax {
			reachableMax = upper
		}

		if prevSum[target] != -1 {
			break
		}
	}

	if prevSum[target] == -1 {
		return nil, false
	}

	out := make([]int, 0)
	for s := target; s > 0; s = int(prevSum[s]) {
		out = append(out, int(picked[s]))
	}

	return out, true
}

func findResidualDeepMatch(list1, list2 []Item) Candidate {
	idx1 := unusedIndicesSorted(list1)
	idx2 := unusedIndicesSorted(list2)

	if len(idx1) == 0 || len(idx2) == 0 {
		return Candidate{}
	}

	var total1, total2 int64
	for _, idx := range idx1 {
		total1 += list1[idx].Value
	}
	for _, idx := range idx2 {
		total2 += list2[idx].Value
	}

	limit := int(min64(total1, total2))
	if limit <= 0 {
		return Candidate{}
	}

	reach1 := buildReachableBitset(list1, idx1, limit)
	reach2 := buildReachableBitset(list2, idx2, limit)

	target := largestCommonReachableSum(reach1, reach2, limit)
	if target <= 0 {
		return Candidate{}
	}

	aIdx, okA := reconstructSubsetForTarget(list1, idx1, target)
	if !okA || len(aIdx) == 0 {
		return Candidate{}
	}

	bIdx, okB := reconstructSubsetForTarget(list2, idx2, target)
	if !okB || len(bIdx) == 0 {
		return Candidate{}
	}

	return Candidate{
		AIdx:    aIdx,
		BIdx:    bIdx,
		Sum:     int64(target),
		Pattern: "residual-dp",
		Ok:      true,
	}
}

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Uso:")
		fmt.Println("  go run conciliar.go lista1.txt lista2.txt")
		return
	}

	list1, err := loadItems(os.Args[1])
	if err != nil {
		fmt.Println("Erro ao carregar lista 1:", err)
		return
	}

	list2, err := loadItems(os.Args[2])
	if err != nil {
		fmt.Println("Erro ao carregar lista 2:", err)
		return
	}

	totalOriginal1 := sumAll(list1)
	totalOriginal2 := sumAll(list2)
	diffOriginal := totalOriginal2 - totalOriginal1

	order1 := buildSortedOrder(list1)
	order2 := buildSortedOrder(list2)

	pairs2 := buildPairsSorted(list2)
	triples2 := buildTriplesSorted(list2)

	maxSum := maxUsefulSum(list2)
	pairBuckets1 := buildPairBuckets(list1, maxSum)

	fmt.Printf("Lista 1: %d itens\n", len(list1))
	fmt.Printf("Lista 2: %d itens\n", len(list2))
	fmt.Printf("Pares da lista 2: %d\n", len(pairs2))
	fmt.Printf("Triplas da lista 2: %d\n", len(triples2))
	fmt.Println("Tentando automaticamente: 1x1, 1x2, 2x1, 2x2, 1x3, 2x3, 3x1, 3x2")
	fmt.Println(strings.Repeat("-", 60))

	var matches []Candidate

	for {
		candidates := []Candidate{
			find1x1(list1, list2, order1, order2),
			find1x2(list1, list2, order1, pairs2),
			find2x1(list1, list2, order2, pairBuckets1),
			find2x2(list1, list2, pairs2, pairBuckets1),
			find1x3(list1, list2, order1, triples2),
			find2x3(list1, list2, triples2, pairBuckets1),
			find3x1(list1, list2, order1, order2, pairBuckets1),
			find3x2(list1, list2, order1, pairs2, pairBuckets1),
		}

		best := Candidate{}
		for _, c := range candidates {
			if better(c, best) {
				best = c
			}
		}

		if !best.Ok {
			break
		}

		markUsed(list1, best.AIdx)
		markUsed(list2, best.BIdx)
		matches = append(matches, best)
	}

	fmt.Println("Fase rápida concluída. Tentando conciliar resíduos...")
	for {
		deep := findResidualDeepMatch(list1, list2)
		if !deep.Ok {
			break
		}

		markUsed(list1, deep.AIdx)
		markUsed(list2, deep.BIdx)
		matches = append(matches, deep)
	}

	saldo1 := sumUnused(list1)
	saldo2 := sumUnused(list2)
	diffResidual := saldo2 - saldo1

	fmt.Println("Processamento concluído.")
	fmt.Printf("Casamentos encontrados: %d\n", len(matches))
	fmt.Printf("Total original lista 1: %s\n", formatBRL(totalOriginal1))
	fmt.Printf("Total original lista 2: %s\n", formatBRL(totalOriginal2))
	fmt.Printf("Diferença original (L2 - L1): %s\n", formatBRL(diffOriginal))
	fmt.Printf("Itens restantes na lista 1: %d | Total: %s\n", countUnused(list1), formatBRL(saldo1))
	fmt.Printf("Itens restantes na lista 2: %d | Total: %s\n", countUnused(list2), formatBRL(saldo2))
	fmt.Printf("Diferença residual (L2 - L1): %s\n", formatBRL(diffResidual))

	if diffOriginal == diffResidual {
		fmt.Println("Diagnóstico: consistente.")
	} else {
		fmt.Println("Diagnóstico: inconsistente.")
	}

	mode := askDisplayMode()

	fmt.Println()
	fmt.Println(strings.Repeat("=", 60))

	switch mode {
	case "1":
		printMatches(matches, list1, list2)
	case "2":
		printUnmatched("Lista 1 - itens sem combinação", list1)
		printUnmatched("Lista 2 - itens sem combinação", list2)
	case "3":
		printMatches(matches, list1, list2)
		fmt.Println()
		fmt.Println(strings.Repeat("=", 60))
		printUnmatched("Lista 1 - itens sem combinação", list1)
		printUnmatched("Lista 2 - itens sem combinação", list2)
	}
}
