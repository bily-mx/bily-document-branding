# Convierte un .docx o .pptx a PDF con Word o PowerPoint (solo Windows con Office instalado).
# Uso: powershell -File scripts/office-a-pdf.ps1 <entrada> [salida.pdf]
param(
  [Parameter(Mandatory = $true)][string]$Entrada,
  [string]$Salida
)
$ErrorActionPreference = 'Stop'
$Entrada = (Resolve-Path $Entrada).Path
if (-not $Salida) { $Salida = [IO.Path]::ChangeExtension($Entrada, '.pdf') }
$Salida = [IO.Path]::GetFullPath($Salida)
$ext = [IO.Path]::GetExtension($Entrada).ToLower()

if ($ext -eq '.docx') {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  try {
    $doc = $word.Documents.Open($Entrada, $false, $true)
    $doc.ExportAsFixedFormat($Salida, 17)   # 17 = PDF
    $doc.Close($false)
  } finally { $word.Quit() }
} elseif ($ext -eq '.pptx') {
  $ppt = New-Object -ComObject PowerPoint.Application
  try {
    $pres = $ppt.Presentations.Open($Entrada, $true, $false, $false)   # solo lectura, sin ventana
    $pres.SaveAs($Salida, 32)   # 32 = PDF
    $pres.Close()
  } finally { $ppt.Quit() }
} else {
  throw "Extension no soportada: $ext (usa .docx o .pptx)"
}
Write-Output "Listo: $Salida"
