param(
    [string]$SourceDir = "W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\ALEATORIOS\magalu_downloads_organizados\manter\meses",
    [string]$OutputPath = "W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\ALEATORIOS\magalu_downloads_organizados\manter\meses\servicos_unificados.xlsx"
)

$ErrorActionPreference = "Stop"

$Atilde = [char]0x00C3
$Ccedilha = [char]0x00C7

$serviceNames = @{
    "co_participation_commission" = "COMISS${Atilde}O DE COPARTICIPA${Ccedilha}${Atilde}O"
    "commission_delivery_service" = "COMISS${Atilde}O SERVI${Ccedilha}O DE ENTREGA"
    "co_participation_cost" = "CUSTO DE COPARTICIPA${Ccedilha}${Atilde}O"
    "delivery_service_cost" = "CUSTO DE SERVI${Ccedilha}O DE ENTREGA"
    "marketplace_services" = "SERVI${Ccedilha}OS DE MARKETPLACE"
    "technology_services" = "SERVI${Ccedilha}OS DE TECNOLOGIA"
    "other_services" = "OUTROS SERVI${Ccedilha}OS"
    "other_fees" = "OUTRAS TARIFAS"
}

function Get-CellExportText {
    param(
        $Cell,
        [string]$Header
    )

    $value = $Cell.Value2
    if ($null -eq $value) {
        return ""
    }

    $text = [string]$Cell.Text
    if ($Header -eq "ID do Pedido") {
        if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
            return ([math]::Round([double]$value)).ToString("0", [System.Globalization.CultureInfo]::InvariantCulture)
        }
        return $text.Trim()
    }

    return $text.Trim()
}

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

$files = Get-ChildItem -Path $SourceDir -File -Filter *.xlsx |
    Where-Object {
        $_.Name -ne [System.IO.Path]::GetFileName($OutputPath) -and
        $_.BaseName -like "service_invoices_*"
    } |
    Sort-Object Name

if (-not $files) {
    throw "Nenhuma planilha service_invoices_*.xlsx foi encontrada em $SourceDir"
}

$excel = New-Object -ComObject Excel.Application
$workbook = $null
$worksheet = $null

try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Servicos"

    $outputRow = 1
    $headerWritten = $false

    foreach ($file in $files) {
        if ($file.Name -notmatch "^service_invoices_(.+?)_\d{4}-\d{2}(?:_\d+)?\.xlsx$") {
            continue
        }

        $serviceKey = $Matches[1]
        $servicePt = $serviceNames[$serviceKey]
        if (-not $servicePt) {
            $servicePt = ($serviceKey -replace "_", " ").ToUpper()
        }

        $sourceWorkbook = $null
        $sourceWorksheet = $null
        $usedRange = $null

        try {
            $sourceWorkbook = $excel.Workbooks.Open($file.FullName, $null, $true)
            $sourceWorksheet = $sourceWorkbook.Worksheets.Item(1)
            $usedRange = $sourceWorksheet.UsedRange

            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count
            if ($rows -lt 2) {
                continue
            }

            $headers = @{}
            for ($col = 1; $col -le $cols; $col++) {
                $header = [string]$sourceWorksheet.Cells.Item(1, $col).Text
                $headers[$header] = $col
            }

            if (-not $headers.ContainsKey("Lojista") -or -not $headers.ContainsKey("ID do Pedido")) {
                continue
            }

            $lojaCol = $headers["Lojista"]
            $pedidoCol = $headers["ID do Pedido"]

            if (-not $headerWritten) {
                for ($col = 1; $col -le $cols; $col++) {
                    $header = [string]$sourceWorksheet.Cells.Item(1, $col).Text
                    if ($col -eq $lojaCol) {
                        $header = "Historico"
                    }
                    $worksheet.Cells.Item($outputRow, $col).Value2 = $header
                    if ($col -eq $pedidoCol) {
                        $worksheet.Columns.Item($col).NumberFormat = "@"
                    }
                }
                $worksheet.Rows.Item($outputRow).Font.Bold = $true
                $outputRow++
                $headerWritten = $true
            }

            for ($row = 2; $row -le $rows; $row++) {
                $pedidoId = Get-CellExportText -Cell $sourceWorksheet.Cells.Item($row, $pedidoCol) -Header "ID do Pedido"
                $rowHasValue = $false

                for ($col = 1; $col -le $cols; $col++) {
                    $header = [string]$sourceWorksheet.Cells.Item(1, $col).Text
                    $sourceCell = $sourceWorksheet.Cells.Item($row, $col)
                    $value = Get-CellExportText -Cell $sourceCell -Header $header

                    if ($value -ne "") {
                        $rowHasValue = $true
                    }

                    $targetCell = $worksheet.Cells.Item($outputRow, $col)
                    if ($col -eq $lojaCol) {
                        $targetCell.Value2 = "$servicePt $pedidoId".Trim()
                    }
                    elseif ($col -eq $pedidoCol) {
                        $targetCell.NumberFormat = "@"
                        $targetCell.Value2 = $value
                    }
                    elseif ($header -eq "Data") {
                        $targetCell.NumberFormat = "@"
                        $targetCell.Value2 = [string]$sourceCell.Text
                    }
                    else {
                        $targetCell.Value2 = $value
                    }
                }

                if ($rowHasValue) {
                    $outputRow++
                }
            }
        }
        finally {
            if ($null -ne $sourceWorkbook) {
                $sourceWorkbook.Close($false)
            }
            Release-ComObject $usedRange
            Release-ComObject $sourceWorksheet
            Release-ComObject $sourceWorkbook
        }
    }

    $worksheet.Columns.AutoFit() | Out-Null

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    $workbook.SaveAs($OutputPath, 51)
    $workbook.Close($true)
    $excel.Quit()
}
finally {
    Release-ComObject $worksheet
    Release-ComObject $workbook
    if ($null -ne $excel) {
        $excel.Quit()
    }
    Release-ComObject $excel
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Planilha criada em: $OutputPath"
