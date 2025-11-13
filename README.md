# Dashboard de Ocupação dos Consultórios

Aplicação Streamlit para consolidar e analisar a ocupação dos consultórios de uma clínica, lendo diretamente as abas **CONSULTÓRIO** e **MÉDICOS** de uma planilha Excel. O dashboard entrega uma visão executiva, rankings de produtividade, análise por plano/aluguel e a agenda completa, além de exportar todo o conteúdo para PDF com o mesmo layout corporativo mostrado na web.

## Principais recursos

- **Importação inteligente do Excel**: detecta automaticamente o cabeçalho das abas de consultório, produtividade e médicos.
- **Filtros globais** por consultório, dia, turno e profissional que são aplicados a todos os componentes relevantes.
- **KPIs de ocupação** com métricas de slots disponíveis, taxa de ocupação e médicos distintos.
- **Rankings e comparativos** baseados nas abas de produtividade (receita, solicitações, cirurgias, exames etc.).
- **Detalhamento por consultório** com cronogramas, heatmaps e distribuição de planos/aluguel.
- **Exportação para PDF** através do `DashboardPDFBuilder`, reutilizando os gráficos Plotly exibidos no navegador.

## Estrutura do projeto

```
.
├── main.py                # Aplicação Streamlit principal
├── app/
│   ├── data/              # Camada de acesso e higienização dos dados Excel
│   │   ├── loader.py      # Funções de leitura das abas e normalização de colunas
│   │   ├── processors.py  # Utilidades (ex.: normalização de nomes, parsing numérico)
│   │   └── facade.py      # `ConsultorioDataFacade` expõe carregamento + filtros
│   ├── export/pdf_builder.py # Montagem do relatório em PDF com fpdf2/kaleido
│   └── services/          # Regras de negócio (ocupação, ranking, timeseries)
├── requirements.txt       # Dependências Python
└── README.md
```

## Pré-requisitos

- Python 3.10+ (recomendado 3.11).
- Ambiente virtual (venv, pipenv, poetry, etc.).
- Arquivo Excel seguindo o layout "ESCALA DOS CONSULTORIOS DEFINITIVO" com as abas:
  - `CONSULTÓRIO` (uma ou várias, ex.: "CONSULTÓRIO 1", "CONSULTÓRIO 2"...)
  - `MÉDICOS` (podem existir múltiplas abas: "MÉDICOS 1", "MÉDICOS 2"...)
  - `PRODUTIVIDADE` (opcional, para liberar KPIs e rankings financeiros)

## Configuração do ambiente

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Como executar

```bash
streamlit run main.py
```

Durante o carregamento o app tentará abrir automaticamente o arquivo definido em `DEFAULT_PATH` (`/mnt/data/ESCALA DOS CONSULTORIOS DEFINITIVO.xlsx`). Caso não exista, utilize o uploader da barra lateral e selecione seu `.xlsx`.

## Fluxo de uso sugerido

1. Abra o dashboard e faça upload da planilha consolidada.
2. Ajuste os filtros globais (consultórios, dias, turnos, médicos) para o cenário desejado.
3. Navegue pelas seções do menu lateral:
   - **Visão Geral**: KPIs, ocupação por sala/dia/turno e médicos com mais turnos.
   - **Ranking**: consolida dados da aba de produtividade (receita total, top médicos, etc.).
   - **Consultórios**: detalha cada sala com gráficos de ocupação e agenda semanal.
   - **Planos & Aluguel**: quebra por planos de saúde ou modalidade de aluguel.
   - **Agenda**: tabela completa com todos os slots filtrados.
4. Clique em **Baixar PDF** para gerar o relatório com os gráficos atuais.

## Boas práticas de desenvolvimento

- Use `streamlit run main.py --server.runOnSave true` durante o desenvolvimento para recarregar automaticamente o app.
- Adicione novas dependências em `requirements.txt` e pinne versões mínimas compatíveis.
- Prefira funções dentro de `app/services/` para regras de negócio reutilizáveis (ex.: cálculos de ocupação, rankings).
- Utilize `ConsultorioDataFacade` para centralizar leitura/filtragem de dados e manter a lógica de I/O isolada da camada de visualização.

## Suporte

Problemas ou sugestões? Abra uma issue descrevendo o cenário e, se possível, anexe um exemplo (anonimizado) do Excel que reproduz o comportamento.
