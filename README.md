# Report Automático — Relatórios Semanais de Marketing Digital

Gerador automático de apresentações `.pptx` para clientes de Ecommerce e Lead.

---

## Instalação

```bash
pip install python-pptx openpyxl
```

---

## Estrutura de pastas

```
report-automatico/
├── main.py               # Ponto de entrada — rode este arquivo
├── data_processor.py     # Leitura e agregação dos dados do xlsx
├── pptx_filler.py        # Preenchimento do template com os valores
├── create_templates.py   # Setup único: cria os templates a partir dos originais
├── clients_config.json   # Configuração dos clientes (metas, pastas)
├── clientes/
│   ├── haux.xlsx         # Planilha do cliente Haux (Ecommerce)
│   └── emerald.xlsx      # Planilha do cliente Emerald (Lead)
├── templates/            # Gerado pelo create_templates.py
│   ├── template_ecommerce.pptx
│   └── template_lead.pptx
└── output/               # Relatórios gerados (fallback quando pasta_drive não existe)
```

---

## Configuração inicial (uma única vez)

1. Coloque as planilhas dos clientes em `clientes/`:
   - `clientes/haux.xlsx` (formato Ecommerce)
   - `clientes/emerald.xlsx` (formato Lead)

2. Gere os templates a partir dos arquivos originais:

```bash
cd report-automatico
python create_templates.py
```

Isso criará `templates/template_ecommerce.pptx` e `templates/template_lead.pptx`.

3. *(Opcional)* Edite `clients_config.json` para ajustar `pasta_drive` e metas mensais.

---

## Uso semanal

```bash
cd report-automatico
python main.py
```

O programa perguntará interativamente:

```
Clientes disponíveis:
  1. haux
  2. emerald
Escolha o cliente: 2

Tipo do cliente:
  1. Ecommerce
  2. Lead
Escolha o tipo: 2

Período analisado (ex: 14/04 a 20/04): 05/12 a 11/12
Período comparado  (ex: 07/04 a 13/04): 28/11 a 04/12

Processando dados...
Gerando apresentação...

Relatório salvo em: output/Emerald_Turism_Heath_21-04.pptx
```

---

## Formato das planilhas

### Ecommerce (`haux.xlsx`)
Colunas obrigatórias (em qualquer ordem, cabeçalho com "DATA"):

| DATA | DIA | VALOR INVESTIDO | FATURAMENTO | ROAS | PEDIDOS | CPS | TAXA DE CONVERSÃO | SESSÕES | TICKET MÉDIO |
|------|-----|-----------------|-------------|------|---------|-----|-------------------|---------|--------------|
| 14/04/2025 | Seg | 500,00 | 3200,00 | 6,4 | 8 | 62,50 | 4,0% | 200 | 400,00 |

### Lead (`emerald.xlsx`)
Colunas obrigatórias:

| DATA | DIA | AÇÃO DO DIA | VALOR INVESTIDO | LEADS | CPL |
|------|-----|-------------|-----------------|-------|-----|
| 05/12/2025 | Sex | Campanha X | 50,00 | 9 | 5,56 |

**Regras de leitura:**
- A linha de cabeçalho é detectada automaticamente pela coluna `DATA`
- Linhas com DATA em maiúsculo (ex: `JANEIRO`) ou `TOTAL` são ignoradas
- ROAS, CPS, TAXA DE CONVERSÃO, CPL e TICKET MÉDIO são recalculados a partir dos totais

---

## Adicionar um novo cliente

1. Adicione a planilha em `clientes/nome_cliente.xlsx`
2. Adicione a entrada em `clients_config.json`:

```json
"novo_cliente": {
  "nome": "Nome Completo do Cliente",
  "pasta_drive": "/caminho/para/pasta/no/drive",
  "metas": {
    "faturamento_mensal": 50000,
    "investimento_mensal": 5000
  }
}
```

---

## Saída

O relatório é salvo automaticamente com o nome `NomeCliente_DD-MM.pptx`:
- Na `pasta_drive` configurada (se existir)
- Em `output/` como fallback
