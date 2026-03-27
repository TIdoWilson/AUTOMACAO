param(
    [string]$DiaTeste = ""
)

Add-Type -AssemblyName PresentationFramework

$agenda = @{
    "07" = "Impostos 2ª Parte"
    "08" = "Provisão e Geração de lote de férias | 13º 2ª Parte"
    "10" = "Processar Adiantamentos"
    "13" = "Relatório de férias, resumo da folha"
    "18" = "Captador de recibos DCTFWeb"
    "20" = "Pro Labore`n1ª Parte`nDoméstica"
    "26" = "Consulta de Empréstimos e Importação no Sistema"
    "27" = "Provisão e Geração de lote de férias | 13º 1ª Parte"
}

$hojeIso = (Get-Date).ToString("yyyy-MM-dd")
$dia = if ([string]::IsNullOrWhiteSpace($DiaTeste)) {
    (Get-Date).ToString("dd")
} else {
    $DiaTeste.PadLeft(2, '0')
}

if (-not $agenda.ContainsKey($dia)) {
    exit 0
}

$estadoPath = Join-Path $PSScriptRoot "alerta_agenda_rh_estado.txt"
$tagHoje = "$hojeIso|$dia"

if (Test-Path $estadoPath) {
    $tagSalva = (Get-Content -Path $estadoPath -Raw -ErrorAction SilentlyContinue).Trim()
    if ($tagSalva -eq $tagHoje) {
        exit 0
    }
}

$titulo = "Agenda Robôs RH - Dia $dia"
$texto = @"
Rotina de hoje:
$($agenda[$dia])

Você já concluiu esta rotina?
- Sim: não alertar mais hoje.
- Não: lembrar novamente na próxima hora.
"@

$resposta = [System.Windows.MessageBox]::Show($texto, $titulo, "YesNo", "Information")

if ($resposta -eq [System.Windows.MessageBoxResult]::Yes) {
    Set-Content -Path $estadoPath -Value $tagHoje -Encoding UTF8
}
