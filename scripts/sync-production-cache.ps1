param(
  [int]$StartYear = 2006,
  [int]$EndYear = (Get-Date).Year,
  [switch]$RefreshCurrentYear
)

$ErrorActionPreference = 'Stop'
$datasetUrl = 'https://datos.gob.ar/api/3/action/package_show?id=produccion-de-petroleo-y-gas-por-pozo'
$accountName = 'capivproxyqe'
$resourceGroup = 'rrhh-portal-rg'
$containerName = 'production-cache'
$cacheDirectory = Join-Path ([IO.Path]::GetTempPath()) 'capiv-production-cache'
$resolvedTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$resolvedCache = [IO.Path]::GetFullPath($cacheDirectory)
if (-not $resolvedCache.StartsWith($resolvedTemp, [StringComparison]::OrdinalIgnoreCase)) {
  throw "El directorio temporal quedó fuera de la carpeta esperada: $resolvedCache"
}
New-Item -ItemType Directory -Path $resolvedCache -Force | Out-Null

function Get-Dataset {
  for ($attempt = 1; $attempt -le 5; $attempt++) {
    try {
      return Invoke-RestMethod -Uri $datasetUrl -TimeoutSec 60
    } catch {
      if ($attempt -eq 5) { throw }
      Start-Sleep -Seconds (2 * $attempt)
    }
  }
}

function Normalize-ResourceName([string]$value) {
  return (($value.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '') -replace '[^a-zA-Z0-9]+', ' ').ToLowerInvariant().Trim()
}

$dataset = Get-Dataset
$candidates = foreach ($resource in $dataset.result.resources) {
  $name = Normalize-ResourceName $resource.name
  if ($resource.format -ne 'CSV') { continue }
  if ($name -notmatch 'produccion de pozos de gas y petroleo') { continue }
  if ($name -match 'ddjj abiertas y cerradas|no convencional') { continue }
  if ($name -notmatch '\b(19|20)\d{2}\b') { continue }
  $year = [int]$Matches[0]
  if ($year -lt $StartYear -or $year -gt $EndYear) { continue }
  [pscustomobject]@{ Year = $year; Url = $resource.url; Modified = $resource.last_modified }
}
$resources = $candidates | Group-Object Year | ForEach-Object {
  $_.Group | Sort-Object Modified -Descending | Select-Object -First 1
} | Sort-Object Year

if (-not $resources) { throw 'No se encontraron recursos anuales de Capítulo IV.' }
$accountKey = az storage account keys list --resource-group $resourceGroup --account-name $accountName --query '[0].value' -o tsv
if (-not $accountKey) { throw 'No se pudo obtener acceso al almacenamiento de CapIV.' }

foreach ($resource in $resources) {
  $blobName = "$($resource.Year).csv"
  $filteredMarker = "filtered/$($resource.Year)/_complete.json"
  $mustRefresh = $RefreshCurrentYear -and $resource.Year -eq (Get-Date).Year
  $previousErrorPreference = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  $existingLength = az storage blob show --account-name $accountName --container-name $containerName --name $blobName --account-key $accountKey --query 'properties.contentLength' -o tsv 2>$null
  $showExitCode = $LASTEXITCODE
  $ErrorActionPreference = $previousErrorPreference
  $rawExists = $showExitCode -eq 0 -and [long]$existingLength -gt 0
  $localFile = Join-Path $resolvedCache $blobName
  if ($rawExists -and -not $mustRefresh) {
    Write-Output "[$($resource.Year)] cache existente ($existingLength bytes)"
  } else {
    Write-Output "[$($resource.Year)] descargando fuente oficial"
    & curl.exe -L --fail --retry 4 --retry-delay 3 --max-time 300 -sS -o $localFile $resource.Url
    if ($LASTEXITCODE -ne 0) { throw "Falló la descarga de $($resource.Year)." }
    $file = Get-Item -LiteralPath $localFile
    if ($file.Length -lt 1000) { throw "El archivo $($resource.Year) es demasiado pequeño ($($file.Length) bytes)." }
    $header = Get-Content -LiteralPath $localFile -Encoding utf8 -TotalCount 1
    if ($header -notmatch 'idareapermisoconcesion' -or $header -notmatch 'idpozo') {
      throw "El archivo $($resource.Year) no tiene el esquema esperado."
    }

    Write-Output "[$($resource.Year)] subiendo $($file.Length) bytes"
    az storage blob upload --account-name $accountName --container-name $containerName --name $blobName --file $localFile --overwrite true --account-key $accountKey --max-connections 8 --no-progress --output none
    if ($LASTEXITCODE -ne 0) { throw "Falló la carga de $($resource.Year)." }
    $rawExists = $true
    Write-Output "[$($resource.Year)] cache bruto listo"
  }

  $previousErrorPreference = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  $markerLength = az storage blob show --account-name $accountName --container-name $containerName --name $filteredMarker --account-key $accountKey --query 'properties.contentLength' -o tsv 2>$null
  $markerExitCode = $LASTEXITCODE
  $ErrorActionPreference = $previousErrorPreference
  if ($markerExitCode -eq 0 -and [long]$markerLength -gt 0 -and -not $mustRefresh) {
    Write-Output "[$($resource.Year)] recortes por area existentes"
    if (Test-Path -LiteralPath $localFile) { Remove-Item -LiteralPath $localFile -Force }
    continue
  }

  if (-not (Test-Path -LiteralPath $localFile)) {
    Write-Output "[$($resource.Year)] recuperando cache bruto"
    az storage blob download --account-name $accountName --container-name $containerName --name $blobName --file $localFile --account-key $accountKey --max-connections 8 --no-progress --output none
    if ($LASTEXITCODE -ne 0) { throw "Falló la recuperación de $($resource.Year)." }
  }

  $filteredDirectory = Join-Path $resolvedCache "filtered-$($resource.Year)"
  $resolvedFiltered = [IO.Path]::GetFullPath($filteredDirectory)
  if (-not $resolvedFiltered.StartsWith($resolvedCache, [StringComparison]::OrdinalIgnoreCase)) {
    throw "El directorio de recortes quedó fuera del cache temporal: $resolvedFiltered"
  }
  if (Test-Path -LiteralPath $resolvedFiltered) { Remove-Item -LiteralPath $resolvedFiltered -Recurse -Force }
  Write-Output "[$($resource.Year)] generando recortes por area"
  node .\proxy\scripts\split-production.js $localFile $resolvedFiltered
  if ($LASTEXITCODE -ne 0) { throw "Falló el recorte de $($resource.Year)." }

  Write-Output "[$($resource.Year)] publicando recortes"
  az storage blob upload-batch --account-name $accountName --destination $containerName --source $resolvedFiltered --destination-path "filtered/$($resource.Year)" --overwrite true --account-key $accountKey --no-progress --output none
  if ($LASTEXITCODE -ne 0) { throw "Falló la publicación de recortes de $($resource.Year)." }
  Remove-Item -LiteralPath $resolvedFiltered -Recurse -Force
  Remove-Item -LiteralPath $localFile -Force
  Write-Output "[$($resource.Year)] listo"
}

Write-Output 'Cache de producción CapIV sincronizado.'
