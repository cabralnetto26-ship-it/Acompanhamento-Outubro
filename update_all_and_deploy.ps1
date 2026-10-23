# ==============================================================================
# Script de Atualização e Publicação no GitHub
# Regional: 847010600 - JOSÉ CABRAL NETO
# ==============================================================================
$ErrorActionPreference = "Stop"

$workspaceDir = "C:\Users\jose.cabral\OneDrive - Hypera\Documentos\Aula IA"
$excelPath = Join-Path $workspaceDir "Acompanhamento 847010600 Outubro26.xlsx"
$copyPath  = Join-Path $env:TEMP "Acompanhamento_temp_$(Get-Random).xlsx"

Write-Output "==> [1/6] Copiando planilha temporariamente para leitura segura..."
Copy-Item $excelPath -Destination $copyPath -Force

$consultantMap = @{
    "847010601" = "Rafael Nascimento"
    "847010602" = "Wellington Morais"
    "847010603" = "Erivaldo Maciel"
    "847010604" = "Leonardo Mazzuco"
    "847164"    = "Leonardo Mazzuco"
    "847010605" = "Isabela Cruccitti"
    "847010606" = "Pricila Moreira"
    "847010607" = "Manuel Artero"
    "847010608" = "Valter Pereira Jr"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Output "==> [2/6] Extraindo dados do Excel (BASE 600, RESUMO e HYMPAR)..."
    $wb = $excel.Workbooks.Open($copyPath, $false, $true)

    # 1. Base 600
    $base = $wb.Sheets.Item("BASE 600")
    $data = $base.UsedRange.Value2
    $rowCount = $data.GetLength(0)

    $clients = @()
    for ($r = 2; $r -le $rowCount; $r++) {
        $cnpj = [string]$data[$r, 1]
        if (-not $cnpj) { continue }

        $cota  = $data[$r, 10]; if ($null -eq $cota -or $cota -lt 0) { $cota = 0 }
        $mdtr  = $data[$r, 11]; if ($null -eq $mdtr -or $mdtr -lt 0) { $mdtr = 0 }
        $sales = $data[$r, 13]; if ($null -eq $sales -or $sales -lt 0) { $sales = 0 }

        $setor = [string]$data[$r, 9]
        $parts = $setor -split " - "
        $id = $parts[0].Trim()
        if ($id -eq "847164") { $id = "847010604" }
        $consultor = $consultantMap[$id]
        if (-not $consultor) { $consultor = $id }

        $clients += [PSCustomObject]@{
            CNPJ      = $cnpj
            Razao     = [string]$data[$r, 4]
            Cidade    = [string]$data[$r, 5]
            UF        = [string]$data[$r, 6]
            Rede      = [string]$data[$r, 7]
            SetorId   = $id
            Consultor = $consultor
            Cota      = [double]$cota
            MDTR      = [double]$mdtr
            Sales     = [double]$sales
            Hympar    = [string]$data[$r, 18]
        }
    }

    # 2. HYMPAR Sheet
    $hymparSheet = $wb.Sheets.Item("HYMPAR")
    $hData = $hymparSheet.UsedRange.Value2
    $hRows = $hData.GetLength(0)

    $rawHympar = @()
    for ($hr = 2; $hr -le $hRows; $hr++) {
        $nomeRaw = [string]$hData[$hr, 1]
        if ($nomeRaw -match "^(8470106\d{2})") {
            $id = $matches[1]
            $consultor = $consultantMap[$id]
            $meta = [int]$hData[$hr, 2]
            $pos  = [int]$hData[$hr, 3]
            $gap  = [int]$hData[$hr, 5]

            $baseCount = ($clients | Where-Object { $_.SetorId -eq $id }).Count
            if ($baseCount -eq 0) { $baseCount = [Math]::Round($meta / 0.8) }

            $rawHympar += [PSCustomObject]@{
                id        = $id
                consultor = $consultor
                base      = $baseCount
                meta      = $meta
                pos       = $pos
                gap       = $gap
            }
        }
    }

    # 2.1 HYMPAR NON-BUYERS (BASE 600 Col 18)
    $strNaoComprou = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0x4E, 0xC3, 0xA3, 0x6F, 0x20, 0x43, 0x6F, 0x6D, 0x70, 0x72, 0x6F, 0x75))
    $strComprouSG  = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0x43, 0x6F, 0x6D, 0x70, 0x72, 0x6F, 0x75, 0x20, 0x53, 0x47, 0x20, 0x28, 0x53, 0x65, 0x6D, 0x20, 0x4D, 0x61, 0x72, 0x63, 0x61, 0x73, 0x29))

    $rawHymparNonBuyers = @()
    foreach ($c in $clients) {
        $isNC = ($c.Hympar -match "N.o Comprou")
        $isSG = ($c.Hympar -eq "Comprou SG")
        if ($isNC -or $isSG) {
            $rawHymparNonBuyers += [PSCustomObject]@{
                cnpj         = $c.CNPJ
                razao        = $c.Razao
                cidade       = $c.Cidade
                uf           = $c.UF
                rede         = $c.Rede
                consultor    = $c.Consultor
                cota         = [Math]::Round($c.Cota, 2)
                mdtr         = [Math]::Round($c.MDTR, 2)
                fat          = [Math]::Round($c.Sales, 2)
                statusHympar = if ($isNC) { $strNaoComprou } else { $strComprouSG }
            }
        }
    }
    $rawHymparNonBuyers = $rawHymparNonBuyers | Sort-Object cota -Descending

    # 3. SETORES
    $setoresHash = @{}
    foreach ($c in $clients) {
        $id = $c.SetorId
        if (-not $setoresHash.ContainsKey($id)) {
            $setoresHash[$id] = @{
                id   = $id
                nome = $c.Consultor
                cota = 0.0
                mdtr = 0.0
                fat  = 0.0
            }
        }
        $setoresHash[$id].cota += $c.Cota
        $setoresHash[$id].mdtr += $c.MDTR
        $setoresHash[$id].fat  += $c.Sales
    }

    $rawDataSetores = @()
    foreach ($s in $setoresHash.Values) {
        $cobMdtr = 0; if ($s.cota -gt 0) { $cobMdtr = ($s.mdtr / $s.cota) * 100 }
        $cobFat  = 0; if ($s.cota -gt 0) { $cobFat  = ($s.fat / $s.cota) * 100 }
        $gap     = $s.cota - $s.fat

        $statusClass = "status-ok"
        $statusStr   = "Na Média"
        if ($cobMdtr -ge 39.0) {
            $statusClass = "status-leader"
            $statusStr   = "Líder"
        } elseif ($cobMdtr -lt 36.0) {
            $statusClass = "status-alert"
            $statusStr   = "Acelerar"
        }

        $rawDataSetores += [PSCustomObject]@{
            id          = $s.id
            nome        = $s.nome
            cota        = [Math]::Round($s.cota, 2)
            mdtr        = [Math]::Round($s.mdtr, 2)
            cobMdtr     = [Math]::Round($cobMdtr, 1)
            fat         = [Math]::Round($s.fat, 2)
            cobFat      = [Math]::Round($cobFat, 1)
            gap         = [Math]::Round($gap, 2)
            status      = $statusStr
            statusClass = $statusClass
        }
    }

    # 4. REDES & CLIENTES COM FATURAMENTO ZERO (SEM COMPRA)
    $targetRedes = @(
        "FARMARCAS", "INDEPENDENTES VIN", "ASSOCIADAS", "HIPER SAUDE", "MASTERFARMA", 
        "REDE SANTA CATARINA", "ECONOMIZE RONDONOPOLIS", "DROGARIA TOTAL", "COMPRE BEM", "DSG FARMA"
    )

    $redesHash = @{}
    $zeroClientsList = @()

    foreach ($c in $clients) {
        $rede = [string]$c.Rede
        if ($rede) { $rede = $rede.Trim().ToUpper() } else { $rede = "OUTROS" }
        if ($rede -match "HIPER SA") { $rede = "HIPER SAUDE" }
        if ($rede -match "ECONOMIZE RONDON") { $rede = "ECONOMIZE RONDONOPOLIS" }

        $isZero = ($c.Sales -le 0.01)
        if ($isZero) {
            $statusCli = "Inativo"
            if ($c.MDTR -gt 0.01) { $statusCli = "Pedido Retido" }

            $zeroClientsList += [PSCustomObject]@{
                cnpj      = $c.CNPJ
                razao     = $c.Razao
                rede      = $rede
                consultor = $c.Consultor
                cota      = [Math]::Round($c.Cota, 2)
                mdtr      = [Math]::Round($c.MDTR, 2)
                status    = $statusCli
            }
        }

        if ($targetRedes -contains $rede) {
            $key = "$rede|$($c.Consultor)"
            if (-not $redesHash.ContainsKey($key)) {
                $redesHash[$key] = @{
                    Rede         = $rede
                    Consultor    = $c.Consultor
                    Lojas        = 0
                    LojasAtivas  = 0
                    LojasZeradas = 0
                    Cota         = 0.0
                    Sales        = 0.0
                    CotaZerada   = 0.0
                }
            }
            $redesHash[$key].Lojas += 1
            $redesHash[$key].Cota  += $c.Cota
            $redesHash[$key].Sales += $c.Sales
            if ($isZero) {
                $redesHash[$key].LojasZeradas += 1
                $redesHash[$key].CotaZerada   += $c.Cota
            } else {
                $redesHash[$key].LojasAtivas  += 1
            }
        }
    }

    $rawRedes = @()
    foreach ($val in $redesHash.Values) {
        $rawRedes += [PSCustomObject]@{
            rede         = $val.Rede
            consultor    = $val.Consultor
            lojas        = $val.Lojas
            lojasAtivas  = $val.LojasAtivas
            lojasZeradas = $val.LojasZeradas
            cota         = [Math]::Round($val.Cota, 2)
            fat          = [Math]::Round($val.Sales, 2)
            cotaZerada   = [Math]::Round($val.CotaZerada, 2)
        }
    }

    $rawZeroClients = $zeroClientsList | Sort-Object Cota -Descending

    # 5. CURVA ABC
    $abcList = @()
    $sortedAll = $clients | Sort-Object Cota -Descending
    foreach ($c in ($sortedAll | Select-Object -First 25)) {
        $abcList += $c
    }
    foreach ($cons in ($consultantMap.Values | Select-Object -Unique)) {
        $topCons = $clients | Where-Object { $_.Consultor -eq $cons } | Sort-Object Cota -Descending | Select-Object -First 5
        foreach ($c in $topCons) {
            if (-not ($abcList | Where-Object { $_.CNPJ -eq $c.CNPJ })) {
                $abcList += $c
            }
        }
    }

    $rawAbc = @()
    foreach ($c in ($abcList | Sort-Object Cota -Descending)) {
        $cob = 0; if ($c.Cota -gt 0) { $cob = ($c.Sales / $c.Cota) * 100 }

        $diag = "Em acompanhamento."
        if ($cob -lt 15.0) { $diag = "Grave represamento." }
        elseif ($cob -lt 30.0) { $diag = "Abaixo da média, requer tração." }
        elseif ($cob -gt 40.0) { $diag = "Bom andamento comercial." }

        $rawAbc += [PSCustomObject]@{
            razao     = $c.Razao
            rede      = $c.Rede
            consultor = $c.Consultor
            cota      = [Math]::Round($c.Cota, 2)
            fat       = [Math]::Round($c.Sales, 2)
            cob       = [Math]::Round($cob, 1)
            diag      = $diag
        }
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    if (Test-Path $copyPath) { Remove-Item $copyPath -Force -ErrorAction SilentlyContinue }
}

Write-Output "==> [3/6] Atualizando arquivos HTML com os novos dados e datas..."

$htmlPath = Join-Path $workspaceDir "grafico_ranking_setores.html"
$indexPath = Join-Path $workspaceDir "index.html"

$html = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)

# Inject JSON blocks
$jsonSetores = $rawDataSetores | ConvertTo-Json -Depth 5 -Compress
$jsonRedes   = $rawRedes       | ConvertTo-Json -Depth 5 -Compress
$jsonHympar  = $rawHympar      | ConvertTo-Json -Depth 5 -Compress
$jsonHymparNB = $rawHymparNonBuyers | ConvertTo-Json -Depth 5 -Compress
$jsonAbc     = $rawAbc         | ConvertTo-Json -Depth 5 -Compress
$jsonZero    = $rawZeroClients | ConvertTo-Json -Depth 5 -Compress

$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawDataSetores = \[[\s\S]*?\];', "const rawDataSetores = $jsonSetores;")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawRedes = \[[\s\S]*?\];', "const rawRedes = $jsonRedes;")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawHympar = \[[\s\S]*?\];', "const rawHympar = $jsonHympar;")
if ($html -match 'const rawHymparNonBuyers = \[[\s\S]*?\];') { $html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawHymparNonBuyers = \[[\s\S]*?\];', "const rawHymparNonBuyers = $jsonHymparNB;") } else { $html = $html.Replace("const rawHympar =", "const rawHymparNonBuyers = $jsonHymparNB;
    const rawHympar =") }
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawAbc = \[[\s\S]*?\];', "const rawAbc = $jsonAbc;")

if ($html -match 'const rawZeroClients = \[[\s\S]*?\];') {
    $html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawZeroClients = \[[\s\S]*?\];', "const rawZeroClients = $jsonZero;")
} else {
    $html = $html.Replace("const rawAbc =", "const rawZeroClients = $jsonZero;`r`n    const rawAbc =")
}

# Update date badge and titles
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'MDTR \d{2}/\d{2} \| Sales \d{2}/\d{2}', 'MDTR 14/10 | Sales 19/10')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'MDTR Realizado \(\d{2}/\d{2}\)', 'MDTR Realizado (14/10)')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'Faturamento Bruto \(\d{2}/\d{2}\)', 'Faturamento Bruto (19/10)')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '\(at.{1,4} \d{2}/\d{2}\)', '(até 14/10)')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '<th>MDTR \d{2}/\d{2} \(R\$\)</th>', '<th>MDTR 14/10 (R$)</th>')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '<th>Fat\. Bruto \d{2}/\d{2} \(R\$\)</th>', '<th>Fat. Bruto 19/10 (R$)</th>')

# Update remaining days simulation
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const fatDiarioAtual = item\.fat / \d+;.*', 'const fatDiarioAtual = item.fat / 12; // 12 dias úteis decorridos até 19/10')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'id="diasRestantesInput" class="filter-input" style="width: 80px;" min="1" max="30" value="\d+"', 'id="diasRestantesInput" class="filter-input" style="width: 80px;" min="1" max="30" value="9"')

# Update filter chips and JS logic thresholds
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '\(&ge;\d+([,\.]\d+)?%\)', '(&ge;39%)')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '\(\d+([,\.]\d+)?%-\d+([,\.]\d+)?%\)', '(36%-39%)')
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '\(&lt;\d+([,\.]\d+)?%\)', '(&lt;36%)')

$html = [System.Text.RegularExpressions.Regex]::Replace($html, "activeStatusFilter === 'high' && d\.cobMdtr < \d+(\.\d+)?", "activeStatusFilter === 'high' && d.cobMdtr < 39")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, "activeStatusFilter === 'mid' && \(d\.cobMdtr < \d+(\.\d+)? \|\| d\.cobMdtr >= \d+(\.\d+)?\)", "activeStatusFilter === 'mid' && (d.cobMdtr < 36 || d.cobMdtr >= 39)")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, "activeStatusFilter === 'low' && d\.cobMdtr >= \d+(\.\d+)?", "activeStatusFilter === 'low' && d.cobMdtr >= 36")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, "const progressClass = item\.cobMdtr >= \d+(\.\d+)? \? 'progress-high' : item\.cobMdtr >= \d+(\.\d+)? \? 'progress-mid' : 'progress-low';", "const progressClass = item.cobMdtr >= 39 ? 'progress-high' : item.cobMdtr >= 36 ? 'progress-mid' : 'progress-low';")

# Save both grafico_ranking_setores.html and index.html in pure UTF-8 without BOM
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $html, $utf8NoBom)
[System.IO.File]::WriteAllText($indexPath, $html, $utf8NoBom)
Write-Output "==> HTMLs salvos com sucesso (grafico_ranking_setores.html e index.html) em UTF-8."

Write-Output "==> [4/6] Preparando git commit..."
if (Test-Path "$env:LOCALAPPDATA\Programs\MinGit\cmd\git.exe") {
    $git = "$env:LOCALAPPDATA\Programs\MinGit\cmd\git.exe"
} else {
    $git = "git"
}

& $git add .
& $git status

$commitMsg = "feat: atualizacao da base de dados e sincronizacao no github - Outubro/2026"
& $git commit -m $commitMsg

Write-Output "==> [5/6] Enviando atualizações para o GitHub (push origin main)..."
& $git push origin main

Write-Output "==> SUCESSO! Dashboard atualizado e publicado no GitHub Pages."
