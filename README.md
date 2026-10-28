# 🤖 RPA Comercial — Automação de Pedidos (WEG & Schulz) ➔ ERP LF

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![Selenium](https://img.shields.io/badge/Selenium-4.20%2B-green?logo=selenium)
![Pandas](https://img.shields.io/badge/Pandas-2.0%2B-darkblue?logo=pandas)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/License-Proprietary-red)

> **Sistema RPA (Robotic Process Automation)** desenvolvido em Python para integração, extração e digitação automatizada de pedidos de compra e cotações entre os portais de clientes B2B (**Mercado Eletrônico / WEG**, **Schulz**) e o ERP corporativo (**LF / Oracle APEX**).

---

## 🎯 Visão Geral

O projeto foi construído para eliminar o trabalho manual repetitivo de entrada de pedidos, mitigando erros de digitação de códigos de produtos (SKUs), cálculos de desconto, conferência de lotes mínimos e regras fiscais complexas.

### 🌟 Destaques de Performance e Confiabilidade:
* **Redução drástica no tempo de digitação**: De ~15 minutos manuais por pedido para menos de 1 minuto por processo automatizado.
* **Inteligência Fiscal Automatizada**: Aplicação de regras de ICMS (TTD 490 para SC, Isenção Pró-Emprego para CNPJs específicos e tratamento CIF/FOB).
* **Gestão Inteligente de Lotes Mínimos**: Detecção de restrições de quantidade no ERP com assistente interativo para adequação do pedido.
* **Separação de Pedidos Mistos**: Identificação automática de itens com finalidades diferentes (Industrialização vs. Uso/Consumo) e divisão em pedidos distintos no ERP com anotação de atendimento parcial.
* **Sincronização de Documentos**: Download do espelho do pedido oficial do cliente e anexo automático em PDF no ERP.

---

## 🧩 Arquitetura dos Módulos

```
Código Novo/
├── WEG Pedidos Novo.py                 # Módulo principal de automação WEG (Mercado Eletrônico ➔ LF)
├── Schulz Pedidos.py                  # Automação de pedidos Schulz Automotiva ➔ LF
├── Schulz Compressores Pedidos.py     # Automação especializada para Schulz Compressores
├── Atualizar Precos Schulz.py         # Sincronização periódica de tabelas de preços
├── RFQ.py                             # Gestão e automação de cotações RFQ
├── historico_rfq.py                   # Banco local/cache de cotações RFQ negociadas
├── importar_rfqs_antigas.py           # Utilitário para carga inicial de histórico RFQ
├── config_weg_lf.json                 # Mapeamentos e regras de negócio WEG
├── config_schulz.json                 # Mapeamentos e parâmetros comerciais Schulz
├── contrato_weg_blumenau.json         # Acordo comercial WEG Blumenau
├── requirements.txt                   # Dependências do projeto
├── .env.example                       # Modelo de variáveis de ambiente
├── Rodar WEG.bat                      # Inicializador rápido WEG (Interface UI)
├── Rodar Schulz.bat                   # Inicializador rápido Schulz
├── Rodar RFQ.bat                      # Inicializador rápido RFQ
└── icones/                            # Ativos visuais e ícones de atalho
```

---

## ⚙️ Principais Funcionalidades

### 1. 🏭 Robô WEG (`WEG Pedidos Novo.py`)
* **Leitura da Fila ME**: Varredura inteligente de transações no Mercado Eletrônico identificando pedidos com status *Não Lida*.
* **Extração Precisa de Dados**: Identificação estrita do número oficial da Ordem de Compra (formato `45XXXXXXXX`), ignorando identificadores internos do portal (Serial ME).
* **Regras de Frete e Transportadoras**:
  * Associação automática de transportadoras homologadas por rota (Linhares, Betim, Blumenau, Gravataí, Joinville).
  * Marcação de *Pedido Complementar* para fretes CIF abaixo do valor mínimo.
* **Conferência de Preços**: Alerta visual proativo caso algum item ultrapasse a margem limite de 35% de desconto em relação à tabela base.

### 2. 🔧 Robô Schulz (`Schulz Pedidos.py` & `Schulz Compressores`)
* Cruzamento das ordens de compra da Schulz com as tabelas de acordo comercial vigentes.
* Resolução automática de códigos de produtos e quantidades.
* Inserção dos pedidos com verificação de itens não localizados.

### 3. 📋 Gestão de Cotações RFQ (`RFQ.py`)
* Histórico consultivo de preços aprovados por RFQ para garantir que pedidos amparados por cotação específica mantenham os preços negociados.

---

## 🚀 Instalação e Configuração

### 1. Pré-requisitos
* **Python 3.10 ou superior** instalado no Windows.
* **Google Chrome** atualizado.
* Acesso à rede com permissão aos portais LF e Mercado Eletrônico.

### 2. Clonando o Repositório
```bash
git clone https://github.com/Guiextreme3/automacao-pedidos-weg-schulz.git
cd automacao-pedidos-weg-schulz
```

### 3. Instalando as Dependências
Recomenda-se a criação de um ambiente virtual:
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 4. Configuração das Credenciais (`.env`)
Copie o modelo de ambiente:
```bash
copy .env.example .env
```
Abra o arquivo `.env` e preencha suas credenciais reais de acesso:
```ini
LF_USER=seu_usuario_lf
LF_PASS=sua_senha_lf

WEG_JVE_USER=seu_login_weg_joinville
WEG_JVE_PASS=sua_senha_weg_joinville

WEG_DI_USER=seu_login_weg_distribuicao
WEG_DI_PASS=sua_senha_weg_distribuicao
```

> ⚠️ **Aviso de Segurança**: O arquivo `.env` e os diretórios de pedidos reais estão cadastrados no `.gitignore` e **nunca** devem ser versionados no Git.

---

## 💻 Como Utilizar

Você pode rodar os robôs diretamente pelo Python ou utilizando os arquivos `.bat` inclusos (que já contam com atalhos de tela cheia e ícones customizados):

### Execução via Linha de Comando:
* **Robô WEG:**
  ```bash
  python "WEG Pedidos Novo.py"
  ```
* **Robô Schulz:**
  ```bash
  python "Schulz Pedidos.py"
  ```
* **Módulo RFQ:**
  ```bash
  python "RFQ.py"
  ```

---

## 🛡️ Tecnologias Utilizadas

* [Python](https://www.python.org/) — Linguagem base da automação.
* [Selenium WebDriver](https://www.selenium.dev/) — Automação de navegadores e controle web.
* [Tkinter](https://docs.python.org/3/library/tkinter.html) — Interface gráfica desktop (GUI) para acompanhamento dos pedidos.
* [Pandas](https://pandas.pydata.org/) & [OpenPyXL](https://openpyxl.readthedocs.io/) — Manipulação de relatórios e tabelas de preços em Excel.
* [PyPDF](https://pypdf.readthedocs.io/) — Extração e tratamento programático de pedidos em PDF.

---

## 👨‍💻 Autor

Desenvolvido por **Guilherme Augusto Santos** ([@Guiextreme3](https://github.com/Guiextreme3)).
