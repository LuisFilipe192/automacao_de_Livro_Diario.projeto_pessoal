# 📘 Automação de Livro Diário

Projeto de automação desenvolvido em **Python** para gerar automaticamente um **Livro Diário em Excel** a partir de arquivos PDF de rateio e arquivos Excel de Movimentos de Caixa.

O sistema lê o PDF e o Excel, extrai as informações relevantes e organiza os dados em uma planilha estruturada com cálculos automáticos de saldo.

Deixando claro que os documentos seguem um padrão e esse script nao vai ler todos os documentos da mesma forma.

---

## 🚀 Objetivo

Automatizar um processo manual recorrente que exigia tempo elevado e estava sujeito a erros humanos, garantindo:

- Maior produtividade  
- Redução de erros  
- Padronização do Livro Diário  
- Geração automática de cálculos  

---

## 🛠 Tecnologias Utilizadas

- Python
- pdfplumber (extração de dados do PDF)  
- openpyxl (manipulação e geração do Excel)  

---

## 📂 Estrutura do Projeto

> Exemplo de estrutura (ajuste para o seu nome de arquivos):
📁 automacao-livro-diario  
├── main.py  
├── RateioPeriodo_Report REF 02-2025.pdf    
├── Movimento_do_Caixa REF 02-2025.xlsx    
├── livro_diario.xlsx     
├── .gitignore  
└── README.md  


---

## ⚙️ Funcionamento do Sistema

O fluxo da aplicação segue a seguinte lógica:

1. Leitura completa do arquivo PDF usando `pdfplumber`  
2. Identificação das linhas que contêm **Data de Rateio**
3. Extração dos campos desejados (ex.: **Guia** e **Emolumento**) nas linhas de movimentação
4. Lê um arquivo Excel base usando `openpyxl` para extrair **Data** ,**Movimento** e **Saída**.
4. Associação de cada movimentação(Entrada e Saída) à **data correta**
5. Organização dos dados em um dicionário/lista
6. Geração do arquivo Excel com estrutura de colunas e formatação (AI assistance)
7. Inserção de fórmulas para cálculo automático do saldo diário

---

## 🧠 Estrutura de Dados

Os dados são organizados internamente no seguinte formato (exemplo):

```python
dados = {
    "01/02/2025": [
        {"guia": "SICASE - XXXX", "emolumento": 123.45},
        {"guia": "SICASE - YYYY", "emolumento": 67.89}
    ],
    "02/02/2025": [
        {"guia": "SICASE - ZZZZ", "emolumento": 50.00}
    ]
}
```

## 🧠 Lógica de Cálculo de Saldo

O saldo segue a regra contábil:

saldo_atual = saldo_anterior + entrada - saida

As fórmulas são inseridas automaticamente na planilha, permitindo atualização dinâmica caso novos valores sejam adicionados manualmente (por exemplo, preenchendo Saídas depois).

## 🤖 Uso de Inteligência Artificial 

A configuração estrutural da **formatação** e **organização dentro do Excel** contou com auxílio de **IA** como ferramenta de apoio técnico.

Entretanto:

- A **lógica do projeto** foi validada e revisada por mim.
- O comportamento das fórmulas foi testado para garantir consistência nos cálculos.
- A estrutura final foi ajustada manualmente para atender à necessidade real do processo.

A Inteligência Artificial foi utilizada como ferramenta de produtividade e suporte técnico, não como substituição de entendimento ou desenvolvimento da lógica do sistema.
