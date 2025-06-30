# 🔁 Sistema de Troca de Técnicos - Claro

Este projeto é uma aplicação com interface gráfica desenvolvida em **Python** usando **CustomTkinter** para automatizar a movimentação de técnicos no sistema via navegador.

---

## 📌 Funcionalidades

- Interface moderna em modo escuro com CustomTkinter.
- Leitura de planilhas Excel com os dados dos contratos.
- Login automático no sistema da Claro via Selenium.
- Troca automática de técnico conforme instruções da planilha.
- Gerenciamento de logins salvos em arquivo `logins.json`.
- Pausar e retomar a execução em tempo real.
- Relatório final gerado em Excel com os resultados.

---

## 📁 Requisitos

- Python 3.10 ou superior
- Google Chrome instalado
- ChromeDriver compatível com sua versão do navegador

---

## 📂 Formato esperado da planilha

A planilha deve conter pelo menos as seguintes colunas (insensível a maiúsculas/minúsculas):

| Coluna Original     | Nome Reconhecido |
|---------------------|------------------|
| contrato            | Contrato         |
| novo tec            | Novo Téc         |
| novo login          | Novo Login       |
| se                  | Se               |

> 📝 Somente as linhas onde a coluna `Se` tiver valor `Trocar` serão processadas.

Outras colunas como `Nome`, `Técnico atual`, `Resultado` serão adicionadas automaticamente no processamento.

---

## ▶️ Como Usar

1. Execute o script principal:

```bash
python nome_do_arquivo.py
```

2. Na interface:
   - Clique em **"Gerenciar Logins"** para adicionar ou remover usuários.
   - Escolha um login no menu.
   - Clique em **"Selecionar Planilha"** e escolha o arquivo Excel com os dados.
   - Escolha a aba da planilha e aguarde a execução.

3. Um arquivo chamado `resultado_troca.xlsx` será gerado ao final com os resultados.

---

## 🧠 Exemplo de estrutura `logins.json`

```json
{
  "Fulano Claro": ["fulano.claro", "senha123"],
  "Ciclano Claro": ["ciclano.claro", "outrasenha"]
}
```

---

