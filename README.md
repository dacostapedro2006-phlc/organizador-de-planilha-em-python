# organizador-de-planilha-em-python
Aqui está o **README pronto para colar no GitHub**, organizado, bonito e totalmente formatado no padrão dos melhores repositórios.

---

# 📘 Organizador de Planilhas em Python

Este projeto é um **processador automático de planilhas escolares**, desenvolvido em Python.
Ele lê um arquivo Excel despadronizado, identifica informações dos alunos usando regras inteligentes e cria arquivos organizados por turma, além de gerar relatórios de inconsistências.

O objetivo é **automatizar a limpeza, verificação e separação dos dados**, facilitando o trabalho administrativo escolar.

---

## 🚀 Funcionalidades Principais

### ✔ Leitura automática de todas as abas do Excel

* Lê e combina todas as planilhas do arquivo.
* Remove linhas vazias e organiza os dados em uma única estrutura.

### ✔ Identificação inteligente de alunos

Reconhece um aluno pela análise da primeira linha:

* Deve conter **nome real** (mais de uma palavra, não iniciando com endereço).
* Não pode ser apenas números.

### ✔ Extração automática de informações

O sistema detecta e separa automaticamente:

| Campo           | Como é encontrado                                |
| --------------- | ------------------------------------------------ |
| **Nome**        | Primeira linha válida                            |
| **Matrícula**   | Regex (6–10 dígitos)                             |
| **Turma**       | Regex + validação em lista de turmas autorizadas |
| **Endereço**    | Linha logo abaixo do nome                        |
| **Responsável** | Texto restante da primeira linha                 |
| **Telefone**    | Procurado nas linhas 2 e 3 usando regex          |

---

## 🛠️ Processamento e Geração de Arquivos

### 📄 1. Arquivo geral consolidado

Gera automaticamente:

```
ALUNOS_TODOS.xlsx
```

### 📁 2. Arquivos separados por turma

Dentro da pasta:

```
/turmas_separadas/
```

Exemplo:

```
10221.xlsx  
11421.xlsx  
22211.xlsx  
```

### ⚠️ 3. Arquivo de alunos sem turma válida

```
TURMA_INVALIDA.xlsx
```

### 📝 4. Relatório de inconsistências

Arquivo:

```
relatorio_erros.txt
```

Contém:

* Alunos sem matrícula
* Alunos sem turma válida

---

## 🧠 Tecnologias Utilizadas

* **Python 3**
* **Pandas** — manipulação de planilhas
* **Regex (re)** — detecção inteligente de padrões
* **Pathlib** — manipulação de caminhos
* Tratamento de exceções e validação de dados

---

## 📦 Estrutura do Projeto

```
/
├── 10221 (1).xlsx            # Arquivo original utilizado como entrada
├── ALUNOS_TODOS.xlsx         # Arquivo consolidado gerado
├── turmas_separadas/         # Arquivos divididos por turma
├── relatorio_erros.txt       # Log dos erros encontrados
└── script.py                 # Código principal do processamento
```

---

## ▶️ Como Executar

1. Instale as dependências:

```bash
pip install pandas
```

2. Execute o script Python:

```bash
python3 script.py
```

3. Os arquivos serão gerados automaticamente na mesma pasta.

---

## 🎯 Objetivo

Este projeto foi criado para **automatizar a organização de planilhas escolares**, reduzindo erros manuais e gerando arquivos padronizados, fáceis de usar por equipes administrativas e pedagógicas.

---

Se quiser, posso também:
✔ adicionar badges do GitHub
✔ criar GIF/prints mostrando o funcionamento
✔ ajudar a subir no GitHub Pages
✔ organizar a estrutura do repositório

Só pedir!
