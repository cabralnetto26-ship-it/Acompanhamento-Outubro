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

# UTF-8 encoded text constants to prevent any PowerShell codepage distortion
$statusLider    = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0x4C, 0xC3, 0xAD, 0x64, 0x65, 0x72)) # Líder
$statusNaMedia  = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0x4E, 0x61, 0x20, 0x4D, 0xC3, 0xA9, 0x64, 0x69, 0x61)) # Na Média
$statusAcelerar = "Acelerar"

$diagAbaixo  = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0x41, 0x62, 0x61, 0x69, 0x78, 0x6F, 0x20, 0x64, 0x61, 0x20, 0x6D, 0xC3, 0xA9, 0x64, 0x69, 0x61, 0x2C, 0x20, 0x72, 0x65, 0x71, 0x75, 0x65, 0x72, 0x20, 0x74, 0x72, 0x61, 0xC3, 0xA7, 0xC3, 0xA3, 0x6F, 0x2E))
$diagBom     = "Bom andamento inicial."
$diagAcomp   = "Em acompanhamento."
$diagGrave   = "Grave represamento."

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
        $cnpj = $data[$r, 1]
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
            CNPJ      = [string]$cnpj
            Razao     = [string]$data[$r, 4]
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
        $statusStr   = $statusNaMedia
        if ($cobMdtr -ge 26.0) {
            $statusClass = "status-leader"
            $statusStr   = $statusLider
        } elseif ($cobMdtr -lt 24.0) {
            $statusClass = "status-alert"
            $statusStr   = $statusAcelerar
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

    # 4. REDES
    $targetRedes = @(
        "FARMARCAS", "INDEPENDENTES VIN", "ASSOCIADAS", "HIPER SAUDE", "MASTERFARMA", 
        "REDE SANTA CATARINA", "ECONOMIZE RONDONOPOLIS", "DROGARIA TOTAL", "COMPRE BEM", "DSG FARMA"
    )

    $redesHash = @{}
    foreach ($c in $clients) {
        if (-not $c.Rede) { continue }
        $rede = $c.Rede.Trim().ToUpper()
        if ($rede -match "HIPER SA") { $rede = "HIPER SAUDE" }
        if ($rede -match "ECONOMIZE RONDON") { $rede = "ECONOMIZE RONDONOPOLIS" }

        if ($targetRedes -contains $rede) {
            $key = "$rede|$($c.Consultor)"
            if (-not $redesHash.ContainsKey($key)) {
                $redesHash[$key] = @{
                    Rede      = $rede
                    Consultor = $c.Consultor
                    Lojas     = 0
                    Cota      = 0.0
                    Sales     = 0.0
                }
            }
            $redesHash[$key].Lojas += 1
            $redesHash[$key].Cota  += $c.Cota
            $redesHash[$key].Sales += $c.Sales
        }
    }

    $rawRedes = @()
    foreach ($val in $redesHash.Values) {
        $rawRedes += [PSCustomObject]@{
            rede      = $val.Rede
            consultor = $val.Consultor
            lojas     = $val.Lojas
            cota      = [Math]::Round($val.Cota, 2)
            fat       = [Math]::Round($val.Sales, 2)
        }
    }

    # 5. CURVA ABC (Top regional + top individual per consultant for active filtering)
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

        $diag = $diagAcomp
        if ($cob -lt 5.0) { $diag = $diagGrave }
        elseif ($cob -lt 15.0) { $diag = $diagAbaixo }
        elseif ($cob -gt 25.0) { $diag = $diagBom }

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
$jsonAbc     = $rawAbc         | ConvertTo-Json -Depth 5 -Compress

$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawDataSetores = \[[\s\S]*?\];', "const rawDataSetores = $jsonSetores;")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawRedes = \[[\s\S]*?\];', "const rawRedes = $jsonRedes;")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawHympar = \[[\s\S]*?\];', "const rawHympar = $jsonHympar;")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const rawAbc = \[[\s\S]*?\];', "const rawAbc = $jsonAbc;")

# Update date badge and titles (using regex and pure ASCII to guarantee matching)
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'MDTR 05/10 \| Sales 08/10', "MDTR 09/10 | Sales 13/10")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'MDTR Realizado \(05/10\)', "MDTR Realizado (09/10)")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'Faturamento Bruto \(08/10\)', "Faturamento Bruto (13/10)")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, '\(at.{1,4} 05/10\)', "(até 09/10)")
$html = $html.Replace("<th>MDTR 05/10 (R$)</th>", "<th>MDTR 09/10 (R$)</th>")
$html = $html.Replace("<th>Fat. Bruto 08/10 (R$)</th>", "<th>Fat. Bruto 13/10 (R$)</th>")

# Update remaining days simulation
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'const fatDiarioAtual = item\.fat / 6;.*', "const fatDiarioAtual = item.fat / 8; // 8 dias úteis decorridos até 13/10")
$html = $html.Replace('id="diasRestantesInput" class="filter-input" style="width: 80px;" min="1" max="30" value="15"', 'id="diasRestantesInput" class="filter-input" style="width: 80px;" min="1" max="30" value="13"')

# Update filter chips and JS logic thresholds
$html = $html.Replace('(&ge;11%)', '(&ge;26%)')
$html = $html.Replace('(9,5%-11%)', '(24%-26%)')
$html = $html.Replace('(&lt;9,5%)', '(&lt;24%)')

$html = $html.Replace("activeStatusFilter === 'high' && d.cobMdtr < 11", "activeStatusFilter === 'high' && d.cobMdtr < 26")
$html = $html.Replace("activeStatusFilter === 'mid' && (d.cobMdtr < 9.5 || d.cobMdtr >= 11)", "activeStatusFilter === 'mid' && (d.cobMdtr < 24 || d.cobMdtr >= 26)")
$html = $html.Replace("activeStatusFilter === 'low' && d.cobMdtr >= 9.5", "activeStatusFilter === 'low' && d.cobMdtr >= 24")
$html = $html.Replace("const progressClass = item.cobMdtr >= 11 ? 'progress-high' : item.cobMdtr >= 9.5 ? 'progress-mid' : 'progress-low';", "const progressClass = item.cobMdtr >= 26 ? 'progress-high' : item.cobMdtr >= 24 ? 'progress-mid' : 'progress-low';")
$html = [System.Text.RegularExpressions.Regex]::Replace($html, 'item\.cobMdtr / 15', 'item.cobMdtr / 38')

# Update chart color thresholds
$html = [System.Text.RegularExpressions.Regex]::Replace($html, "d\.cobMdtr >= 12 \? 'rgba\(16, 185, 129, 0\.85\)' :[\s\r\n]*d\.cobMdtr >= 10 \? 'rgba\(59, 130, 246, 0\.85\)' :[\s\r\n]*d\.cobMdtr >= 9  \? 'rgba\(6, 182, 212, 0\.85\)' :", "d.cobMdtr >= 26 ? 'rgba(16, 185, 129, 0.85)' :`r`n                d.cobMdtr >= 24 ? 'rgba(59, 130, 246, 0.85)' :`r`n                d.cobMdtr >= 22 ? 'rgba(6, 182, 212, 0.85)' :")

# Save both grafico_ranking_setores.html and index.html in pure UTF-8 without BOM
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $html, $utf8NoBom)
[System.IO.File]::WriteAllText($indexPath, $html, $utf8NoBom)
Write-Output "==> HTMLs salvos com sucesso (grafico_ranking_setores.html e index.html) em UTF-8."

Write-Output "==> [4/6] Atualizando README.md..."
$readmePath = Join-Path $workspaceDir "README.md"
$readme = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)

$readmeContent = @"
# 📊 Acompanhamento de Setores, Cota e MDTR — Outubro/2026
**Regional:** 847010600 - JOSÉ CABRAL NETO  
**Empresa:** Hypera Pharma  
**Painel Online (GitHub Pages):** [https://cabralnetto26-ship-it.github.io/Acompanhamento-Outubro/](https://cabralnetto26-ship-it.github.io/Acompanhamento-Outubro/)

---

## 🎯 Visão Geral do Projeto

Este repositório contém a análise comercial consolidada e o **dashboard executivo interativo** de acompanhamento da equipe comercial (8 setores), cruzando **Cota Total**, **MDTR** (posição 09/10/2026) e **Faturamento Bruto Realizado** (posição 13/10/2026).

### 📈 Indicadores Principais da Regional

| Indicador | Valor Consolidado | % Atingido | Observação |
| :--- | :---: | :---: | :--- |
| **Cota Total da Regional** | **R$ 8.071.973,31** | 100% | Cota de Outubro/2026 |
| **MDTR Realizado (até 09/10)** | **R$ 2.120.257,37** | **26,3%** | Ritmo acelerado de entrada de pedidos |
| **Faturamento Bruto (até 13/10)** | **R$ 1.428.579,15** | **17,7%** | Pedidos faturados |
| **GAP Restante para Cota** | **-R$ 6.643.394,16** | 82,3% | Volume a faturar em 13 dias úteis restantes |
| **Meta Intermediária** | **R$ 4.893.633,82** | **60,6%** | Alvo intermediário da regional |
| **GAP Meta Intermediária** | **-R$ 3.465.054,67** | - | Volume para o marco intermediário |

---

## 🏆 Ranking dos Setores (Cota vs MDTR)

| Pos. | Setor / Consultor | Cota (R$) | MDTR 09/10 (R$) | Cob. MDTR (%) | Fat. Bruto 13/10 (R$) | Cob. Fat (%) | Status |
| :---: | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **1º** | **847010607 - Manuel Artero** | R$ 1.090.998,13 | R$ 380.814,07 | **34,9%** | R$ 201.177,41 | 18,4% | 🌟 Líder Regional |
| **2º** | **847010606 - Pricila Moreira** | R$ 556.481,27 | R$ 149.131,48 | **26,8%** | R$ 88.265,37 | 15,9% | ✅ Alta Cobertura |
| **3º** | **847010608 - Valter Pereira Jr** | R$ 1.085.968,45 | R$ 289.438,38 | **26,7%** | R$ 148.632,49 | 13,7% | ✅ Alta Cobertura |
| **4º** | **847010605 - Isabela Cruccitti** | R$ 1.363.710,73 | R$ 351.041,23 | **25,7%** | R$ 326.031,50 | 23,9% | ⚡ Na Média |
| **5º** | **847010601 - Rafael Nascimento** | R$ 942.882,94 | R$ 232.405,16 | **24,6%** | R$ 133.047,97 | 14,1% | ⚡ Na Média |
| **6º** | **847010603 - Erivaldo Maciel** | R$ 1.122.794,36 | R$ 270.638,17 | **24,1%** | R$ 182.801,87 | 16,3% | ⚡ Na Média |
| **7º** | **847010604 - Leonardo Mazzuco** | R$ 1.060.357,11 | R$ 252.860,24 | **23,8%** | R$ 211.890,17 | 20,0% | ⚠️ Acelerar |
| **8º** | **847010602 - Wellington Morais** | R$ 848.780,32 | R$ 193.928,64 | **22,8%** | R$ 136.732,37 | 16,1% | ⚠️ Acelerar |

---

## 🎯 Cobertura de Carteira e Positivação (HYMPAR)

- **Clientes Totais na Base:** 1.097
- **Meta de Positivação (80%):** 878
- **Positivados Realizados:** 579 (66,0% de atingimento da meta)
- **Saldo / Gap Restante:** -299 clientes a positivar

---

## 💻 Painel Interativo (`index.html`)

O dashboard está hospedado no GitHub Pages e conta com recursos avançados:
- **Filtros Editáveis:** busca em tempo real por nome/código, filtro por nível de cobertura, e botões para ligar/desligar setores individuais.
- **Simulador de Metas Dinâmico:** ajuste o número de dias úteis restantes e o alvo da cota para ver a meta diária e a taxa de aceleração recalculadas na hora.
- **Gráficos com Chart.js:** visualização horizontal de ranking e barras agrupadas de Cota vs MDTR vs Faturado.
- **Diagnóstico por Rede & Curva ABC:** tabelas dinâmicas que filtram conforme a seleção dos consultores.

---

## 📁 Estrutura de Arquivos

- `index.html`: Dashboard interativo publicado no GitHub Pages.
- `grafico_ranking_setores.html`: Cópia local do painel interativo.
- `ANALISE_SETORES_COTA_MDTR_OUTUBRO2026.md`: Relatório executivo e tático completo.
- `Acompanhamento 847010600 Outubro26.xlsx`: Planilha com a base de dados original.
- `update_all_and_deploy.ps1`: Script automatizado de atualização e push ao GitHub.
"@

[System.IO.File]::WriteAllText($readmePath, $readmeContent, $utf8NoBom)
Write-Output "==> README.md atualizado."

Write-Output "==> [5/6] Preparando git commit..."
$git = "$env:LOCALAPPDATA\Programs\MinGit\cmd\git.exe"

& $git add .
& $git status

$commitMsg = "feat: atualizacao de dados comerciais - MDTR 09/10 e Sales 13/10"
& $git commit -m $commitMsg

Write-Output "==> [6/6] Enviando atualizações para o GitHub (push origin main)..."
& $git push origin main

Write-Output "==> SUCESSO! Dashboard atualizado e publicado no GitHub Pages."
