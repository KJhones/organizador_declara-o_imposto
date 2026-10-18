# 📊 Controle de Imposto de Renda em Excel

Ferramenta construída **inteiramente no Excel** para organizar e reunir, em um único lugar, as informações necessárias para a declaração de Imposto de Renda de Pessoa Física (IRPF). O projeto foi desenvolvido como desafio da [DIO](https://www.dio.me/) e aplica menus de navegação, validação de dados, fórmulas, formatação condicional e links rápidos, sem macros e sem dependências externas.

> ⚠️ **Aviso:** a planilha foi desenvolvida como parte do desafio da DIO e serve para **organizar dados e estimar valores**. Ela não substitui o programa oficial da Receita Federal. As regras, limites e tabelas do IR mudam todo ano: confira sempre os parâmetros da aba `Config` antes de usar.

---

## 🎯 Objetivo

Criar um **agregador de dados** para a declaração de imposto de renda em que o usuário:

- registre rendimentos, despesas dedutíveis, bens, dívidas e dependentes;
- tenha os dados **validados automaticamente** na digitação;
- navegue entre as áreas por um **menu** com links;
- acompanhe um **resumo** com totais, deduções e imposto estimado;
- controle quais **documentos** ainda estão pendentes.

## ✨ Funcionalidades

| Recurso | Descrição |
|---|---|
| **Menu de navegação** | Página inicial com botões (hiperlinks internos) para todas as abas e botão de retorno em cada uma |
| **Links rápidos** | Atalhos para Meu Imposto de Renda, e-CAC, consulta de CPF, Meu INSS e site da Receita Federal |
| **Validação de dados** | Datas dentro do ano-calendário, valores não negativos, CPF/CNPJ com 11 ou 14 dígitos e listas suspensas |
| **Classificação automática** | O tipo de rendimento define sozinho se ele é tributável, isento ou de tributação exclusiva |
| **Resumo consolidado** | Soma por categoria com `SUMIFS`, aplica limites legais e calcula o imposto estimado |
| **Alertas visuais** | Linhas incompletas em laranja, reembolso maior que o gasto em vermelho e contadores de pendências |
| **Checklist de documentos** | Status em lista suspensa com cores (Pendente, Em andamento, Concluído) e percentual de conclusão |
| **Configuração central** | Ano-base, limites e tabela do IR em uma única aba, usados por todas as fórmulas |

## 🗂️ Estrutura do arquivo

O arquivo `Controle_Imposto_de_Renda.xlsx` tem 9 abas:

| Aba | Função |
|---|---|
| `Menu` | Navegação, links rápidos e guia de uso |
| `Resumo` | Totais, deduções, imposto estimado, patrimônio e alertas |
| `Rendimentos` | Cadastro de rendimentos recebidos (200 linhas) |
| `Despesas` | Cadastro de despesas dedutíveis (200 linhas) |
| `Bens_Direitos` | Situação dos bens em 31/12 do ano anterior e do ano-calendário (200 linhas) |
| `Dividas` | Dívidas e ônus reais em 31/12 (200 linhas) |
| `Dependentes` | Cadastro de dependentes (56 linhas) |
| `Checklist` | Documentos necessários e status de cada um |
| `Config` | Parâmetros editáveis, tabela progressiva e listas das validações |

### Colunas das abas de cadastro

**Rendimentos:** Data · Fonte pagadora · CPF/CNPJ · Tipo de rendimento · Valor bruto · IR retido · *Classificação (auto)* · *Mês (auto)* · Comprovante · Observações

**Despesas:** Data · Prestador/beneficiário · CPF/CNPJ · Categoria · Valor pago · Reembolsado · *Valor dedutível (auto)* · Comprovante · Observações

**Bens_Direitos / Dividas:** Grupo ou credor · Discriminação · Situação em 31/12 do ano anterior · Situação em 31/12 do ano-calendário · *Variação (auto)* · Observações

**Dependentes:** Nome · CPF · Parentesco · Data de nascimento · *Idade em 31/12 (auto)* · Observações

> As colunas em *itálico* são calculadas por fórmula (fundo cinza). As demais são de preenchimento manual (fonte azul).

## 🔧 Detalhamento técnico

### 1. Nomes definidos (Named Ranges)

Os parâmetros e as listas são referenciados por nomes, o que deixa as fórmulas legíveis e evita endereços soltos:

| Nome | Referência | Uso |
|---|---|---|
| `AnoBase` | `Config!B3` | Ano-calendário da declaração |
| `LimEduc` | `Config!B4` | Limite anual de dedução de educação por pessoa |
| `DedDep` | `Config!B5` | Dedução anual por dependente |
| `LimPGBL` | `Config!B6` | Limite do PGBL como percentual da renda tributável |
| `ListaTipos` | `Config!A10:A19` | Tipos de rendimento |
| `ListaCategorias` | `Config!D10:D14` | Categorias de despesa |
| `ListaParentesco` | `Config!F10:F15` | Graus de parentesco |
| `ListaStatus` | `Config!H10:H12` | Status do checklist |

### 2. Validação de dados

| Campo | Tipo de validação | Regra |
|---|---|---|
| Datas (Rendimentos e Despesas) | Data | Entre `DATE(AnoBase,1,1)` e `DATE(AnoBase,12,31)` |
| Valores em R$ | Decimal | Maior ou igual a zero |
| CPF/CNPJ | Personalizada | `=OR(LEN(C5)=11, LEN(C5)=14)` (somente números) |
| CPF do dependente | Personalizada | `=LEN(B5)=11` |
| Nascimento do dependente | Data | Não pode ser posterior ao ano-calendário |
| Tipo, categoria, parentesco, status | Lista | Aponta para os nomes definidos (`ListaTipos` etc.) |

Todas exibem mensagem de erro em português e, quando aplicável, mensagem de orientação ao selecionar a célula.

### 3. Fórmulas principais

**Classificação automática do rendimento** (`Rendimentos!G`)

```excel
=IF(D5="","",IFERROR(INDEX(Config!$B$10:$B$19,MATCH(D5,Config!$A$10:$A$19,0)),""))
```

**Valor dedutível da despesa** (`Despesas!G`): evita resultado negativo

```excel
=IF(E5="","",MAX(0,E5-N(F5)))
```

**Idade em 31/12** (`Dependentes!E`)

```excel
=IF(D5="","",DATEDIF(D5,DATE(AnoBase,12,31),"Y"))
```

**Totais por categoria** (`Resumo`)

```excel
=SUMIFS(Rendimentos!$E$5:$E$204, Rendimentos!$G$5:$G$204, "Tributável")
```

**Limite do PGBL**

```excel
=MIN(SUMIFS(Despesas!$G$5:$G$204, Despesas!$D$5:$D$204, "PGBL"), LimPGBL*B6)
```

**Limite de educação** (por pessoa: titular + dependentes)

```excel
=MIN(SUMIFS(Despesas!$G$5:$G$204, Despesas!$D$5:$D$204, "Educação"), LimEduc*(1+B18))
```

**Imposto pela tabela progressiva** (base × alíquota − parcela a deduzir)

```excel
=MAX(0, B22*INDEX(Config!$B$24:$B$28, MATCH(B22, Config!$A$24:$A$28, 1))
        - INDEX(Config!$C$24:$C$28, MATCH(B22, Config!$A$24:$A$28, 1)))
```

O `MATCH(..., 1)` faz busca aproximada na coluna de limites inferiores e localiza a faixa correta; o `INDEX` retorna a alíquota e a parcela a deduzir daquela faixa.

**Cálculo do saldo**

```
Base de cálculo = MAX(0, Rendimentos tributáveis − Total de deduções)
Saldo           = Imposto devido − Imposto retido na fonte
                  (positivo = a pagar | negativo = a restituir)
```

### 4. Formatação condicional

| Local | Condição | Efeito |
|---|---|---|
| Rendimentos e Despesas | Há valor, mas faltam fonte, CPF/CNPJ ou tipo/categoria | Linha em laranja |
| Despesas (coluna F) | Reembolso maior que o valor pago | Célula em vermelho |
| Checklist (coluna C) | Concluído / Em andamento / Pendente | Verde / amarelo / vermelho |
| Resumo (alertas) | Contador maior que zero / igual a zero | Vermelho / verde |

### 5. Navegação

Os botões usam a função `HYPERLINK`, que funciona tanto para abas internas quanto para URLs:

```excel
=HYPERLINK("#'Rendimentos'!A1", "► Rendimentos")        ← aba interna
=HYPERLINK("https://cav.receita.fazenda.gov.br", "↗ Portal e-CAC")   ← link externo
```

### 6. Classificação dos rendimentos (padrão)

| Tipo de rendimento | Classificação |
|---|---|
| Salário / trabalho assalariado, aposentadoria/pensão, aluguel de pessoa física, trabalho autônomo | Tributável |
| 13º salário, PLR, aplicações financeiras (CDB, renda fixa) | Exclusiva/definitiva |
| Poupança/LCI/LCA, lucros e dividendos, FGTS/indenizações | Isento |

A tabela está em `Config!A10:B19` e pode ser editada. Cada alteração se reflete automaticamente nas validações e no Resumo.

## 🚀 Como usar

1. **Baixe** o arquivo `Controle_Imposto_de_Renda.xlsx` e abra no Excel.
2. Na aba `Config`, confira o **ano-calendário** e os **limites** (células amarelas).
3. Preencha as abas de cadastro (`Rendimentos`, `Despesas`, `Bens_Direitos`, `Dividas`, `Dependentes`) nas colunas de **fonte azul**.
4. **Apague as linhas de exemplo** (linha 5 de cada aba de cadastro) antes de usar, pois elas entram nos totais.
5. Acompanhe os documentos na aba `Checklist`.
6. Consulte a aba `Resumo` para ver os totais, os alertas e o saldo estimado.
7. Use os **links rápidos** do `Menu` para acessar os serviços oficiais.

**Dica:** digite CPF e CNPJ **somente com números**. As colunas já estão formatadas como texto para preservar zeros à esquerda.

## 🔄 Como atualizar para um novo ano

1. Em `Config!B3`, altere o **ano-calendário**.
2. Atualize **limite de educação**, **dedução por dependente** e **limite de PGBL** (`B4` a `B6`).
3. Substitua a **tabela progressiva** (`Config!A24:C28`) pelos valores vigentes.
4. Revise a **classificação dos tipos de rendimento** (`Config!A10:B19`), pois isenções e tributações podem mudar.
5. Limpe os dados do ano anterior nas abas de cadastro.

## 🎨 Convenções visuais

| Cor | Significado |
|---|---|
| Fonte azul | Entrada manual do usuário |
| Fonte preta | Fórmula ou texto fixo |
| Fundo cinza | Coluna calculada automaticamente |
| Fundo amarelo | Parâmetro-chave editável (aba `Config`) |
| Fundo laranja | Linha com dados incompletos |

## ⚠️ Limitações conhecidas

- **É uma estimativa.** A tabela progressiva e os limites são parâmetros editáveis informados pelo autor e precisam ser confirmados na Receita Federal a cada ano.
- A tabela anual é uma **simplificação** (12 vezes a tabela mensal) e não considera todas as situações do programa oficial, como o desconto simplificado, a comparação entre modelos de declaração e a redução específica de cada ano.
- O limite de educação é aplicado como "limite por pessoa × (titular + dependentes)", o que é uma aproximação.
- Não há controle de rendimentos recebidos de pessoa jurídica no exterior, ganhos de capital, renda variável ou atividade rural.
- As abas de cadastro comportam 200 linhas (56 em `Dependentes`). Para ampliar, copie as fórmulas e as validações para as linhas seguintes.
- A validação de CPF/CNPJ verifica apenas a **quantidade de dígitos**, não os dígitos verificadores.

## 🛠️ Tecnologias e compatibilidade

- **Microsoft Excel** (formato `.xlsx`), sem macros (VBA)
- Funções usadas: `SUMIFS`, `INDEX`, `MATCH`, `IFERROR`, `HYPERLINK`, `COUNTIF`, `COUNTA`, `SUMPRODUCT`, `DATEDIF`, `MIN`, `MAX`, `N`
- Recursos: nomes definidos, validação de dados (lista, data, decimal e personalizada), formatação condicional, filtros e painéis congelados
- Testado também no LibreOffice Calc, com recálculo sem erros de fórmula

## 💡 Aprendizados

- Uso de **nomes definidos** para tornar fórmulas legíveis e fáceis de manter.
- Combinação de **`INDEX` + `MATCH`** para buscas e para faixas progressivas.
- Camada de **validação** que impede dados inválidos já na digitação.
- Separação clara entre **entrada**, **cálculo** e **resumo**.
- Uso de `HYPERLINK` para criar um **menu de navegação** sem macros.
- Documentação de premissas e limitações como parte do projeto.
