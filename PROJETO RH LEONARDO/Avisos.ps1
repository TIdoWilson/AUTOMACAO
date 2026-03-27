Add-Type -AssemblyName PresentationFramework

$agenda = @{
    "07" = "Impostos 2ª Parte"
    "08" = "Provisão de férias | 13º 2ª Parte"
    "10" = "Adiantamento"
    "20" = "Pro Labore`n1ª Parte`nDoméstica"
    "26" = "Empréstimos DET"
    "27" = "Provisão de férias | 13º 1ª Parte"
}

$dia = (Get-Date).ToString("dd")

if ($agenda.ContainsKey($dia)) {
    $titulo = "Agenda Robôs RH - Dia $dia"
    $texto = $agenda[$dia]
    [System.Windows.MessageBox]::Show($texto, $titulo, "OK", "Information") | Out-Null
}
