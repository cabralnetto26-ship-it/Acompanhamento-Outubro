# 📚 Contexto Completo do Projeto – Acompanhamento de Setores, Cota e MDTR (Outubro/2026)

## 📂 Estrutura de Arquivos

| Arquivo | Descrição |
|---|---|
| `README.md` | Visão geral, indicadores principais e descrição do painel interativo. |
| `index.html` | Dashboard principal publicado no GitHub Pages (`https://cabralnetto26-ship-it.github.io/Acompanhamento-Outubro/`). |
| `grafico_ranking_setores.html` | Cópia local do painel interativo, contém filtros editáveis, simulador de metas e gráficos Chart.js. |
| `ANALISE_SETORES_COTA_MDTR_OUTUBRO2026.md` | Relatório executivo detalhado com análise de setores, métricas de Cota, MDTR e diagnóstico por Cliente e Rede (BASE 600). |
| `Acompanhamento 847010600 Outubro26.xlsx` | Fonte de dados original – planilha com informações de vendas, metas e cobertura. |
| `.gitignore` | Configurações para ignorar arquivos temporários e de build. |
| `setup_git.ps1` | Script que instala MinGit, cria alias `git` no PowerShell e configura credenciais para push ao GitHub. |

## 🎯 Objetivo do Projeto

Fornecer um **painel executivo interativo** que consolide os indicadores comerciais da regional **847010600 – José Cabral Neto** para o mês de **Outubro/2026**.  Ele cruza:

- **Cota Total da Regional**
- **MDTR Realizado (até 05/10/2026)**
- **Faturamento Bruto Realizado (até 08/10/2026)**
- **Diagnóstico por Cliente e Rede (BASE 600)** – análise adicional solicitada pelo usuário.

O dashboard permite que usuários explorem os dados por setor, consultor, cobertura e, através de filtros editáveis, ajustem parâmetros como **dias úteis restantes** e **meta intermediária** para visualizar o ritmo de desempenho em tempo real.

## 📊 Principais Métricas (visíveis no README)

| Indicador | Valor Consolidado | % Atingido | Observação |
|---|---|---|---|
| **Cota Total da Regional** | **R$ 8.071.973,31** | 100% | Cota de Outubro/2026 |
| **MDTR Realizado (até 05/10)** | **R$ 836.984,11** | **10,4%** | Ritmo inicial de entrada de pedidos |
| **Faturamento Bruto (até 08/10)** | **R$ 642.363,76** | **8,0%** | Pedidos faturados |
| **GAP Restante para Cota** | **-R$ 7.429.609,55** | 92,0% | Volume a faturar em 15 dias úteis restantes |
| **Meta Intermediária** | **R$ 4.893.633,82** | **60,6%** | Alvo intermediário da regional |

## 📈 Ranking dos Setores (Cota vs MDTR)

| Pos. | Setor / Consultor | Cota (R$) | MDTR 05/10 (R$) | Cob. MDTR (%) | Fat. Bruto 08/10 (R$) | Cob. Fat (%) | Status |
|---|---|---|---|---|---|---|---|
| **1º** | **847010605 - Isabela Cruccitti** | 1.363.710,73 | 188.273,82 | **13,8%** | 159.685,52 | 11,7% | 🌟 Líder Regional |
| **2º** | **847010607 - Manuel Artero** | 1.090.998,13 | 129.573,32 | **11,9%** | 94.177,28 | 8,6% | ✅ Alta Cobertura |
| **3º** | **847010606 - Pricila Moreira** | 556.481,27 | 57.250,91 | **10,3%** | 38.809,28 | 7,0% | ✅ Na Média |
| **4º** | **847010608 - Valter Pereira Jr** | 1.085.968,45 | 104.638,22 | **9,6%** | 47.842,79 | 4,4% | ⚠️ Acelerar Faturamento |
| **5º** | **847010601 - Rafael Nascimento** | 942.882,94 | 88.775,39 | **9,4%** | 55.371,58 | 5,9% | ⚠️ Acelerar |
| **6º** | **847010603 - Erivaldo Maciel** | 1.122.794,36 | 104.241,41 | **9,3%** | 92.898,10 | 8,3% | ⚠️ Acelerar |
| **7º** | **847010602 - Wellington Morais** | 848.780,32 | 73.902,85 | **8,7%** | 65.384,84 | 7,7% | 🛑 Atenção |
| **8º** | **847010604 - Leonardo Mazzuco** | 1.060.357,11 | 90.328,19 | **8,5%** | 88.194,37 | 8,3% | 🛑 Atenção |

## 🛠️ Tecnologias Utilizadas

- **HTML5** – estrutura da página.
- **CSS (Vanilla)** – layout responsivo, cores premium (gradientes, tipografia Inter via Google Fonts) e micro‑animações.
- **JavaScript** – lógica de filtros, simulador de metas, obtenção de dados via `fetch` (arquivo JSON gerado a partir da planilha).
- **Chart.js** – gráficos horizontais de ranking e barras agrupadas.
- **Git & GitHub Pages** – controle de versão e publicação automática.
- **PowerShell (MinGit)** – script `setup_git.ps1` instala Git e cria alias para usar no ambiente Windows.

## 🔧 Como Executar / Testar Localmente

1. **Clonar o repositório**
   ```powershell
   git clone https://github.com/cabralnetto26-ship-it/Acompanhamento-Outubro.git
   cd Acompanhamento-Outubro
   ```
2. **Abrir o dashboard** – abra `index.html` ou `grafico_ranking_setores.html` em um navegador moderno.
3. **Editar filtros** – os campos de busca, cobertura mínima e botões de toggle são responsivos e atualizam os gráficos em tempo real.
4. **Ajustar meta** – altere “Dias úteis restantes” e “Meta intermediária” na seção *Simulador de Metas* para recalcular ritmo diário e taxa de aceleração.

## 📂 Dados de Diagnóstico por Cliente e Rede (BASE 600)

A análise adicional foi inserida ao final do arquivo `grafico_ranking_setores.html` nas abas **Diagnóstico – Cliente** e **Diagnóstico – Rede**. Cada aba contém:

- Tabela detalhada com os principais indicadores por cliente/por rede (Cota, MDTR, % Cobertura, Faturamento).
- Gráficos de barras que comparam a performance atual com a meta de BASE 600.
- Filtros para seleção de cliente ou rede específicos.

## 📦 Deploy no GitHub Pages

- O repositório está configurado para servir a pasta raiz via GitHub Pages.
- O arquivo `index.html` referencia recursos relativos; ao fazer *push* para a branch `main`, o site é atualizado automaticamente.

## 📋 Próximos Passos sugeridos

- Automatizar a geração dos JSONs a partir da planilha Excel usando um script Python ou PowerShell.
- Implementar exportação de dados filtrados para CSV.
- Adicionar autenticação básica para restringir acesso ao dashboard interno.

---

*Este documento foi gerado automaticamente para consolidar todo o contexto do projeto, facilitando a compreensão, manutenção e extensibilidade por parte de desenvolvedores e stakeholders.*
