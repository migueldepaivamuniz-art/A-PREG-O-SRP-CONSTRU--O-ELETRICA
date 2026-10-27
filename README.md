# Sistema de Pesquisa de Preços de Mercado — Pregão SRP

Sistema corporativo desktop para automação de pesquisa de preços de mercado, auditoria por inteligência artificial e geração de comprovantes oficiais com carimbo monocromático de autenticidade, em estrita conformidade com a **Lei Federal nº 14.133/2021** (Nova Lei de Licitações) e Instrução Normativa SEGES/ME.

---

## 🚀 Como Executar em 1 Clique (Simples e Direto)

Basta dar um **duplo clique** em qualquer um dos arquivos abaixo na pasta do projeto:

- **`INICIAR_SISTEMA.bat`** ou **`EXECUTAR.bat`**

O iniciador automático:
1. Detecta o Python automaticamente no seu Windows.
2. Cria o arquivo `.env` de configuração se ainda não existir.
3. Instala as dependências necessárias de forma automática na primeira execução.
4. Abre a interface gráfica do sistema.

---

## 🛠️ Instalação Manual (Opcional)

Se preferir utilizar a linha de comando:

```bash
# 1. Instalar dependências
pip install -r requirements.txt

# 2. Instalar o navegador do robô de coleta
playwright install chromium

# 3. Iniciar o aplicativo
python app.py
```

---

## 🔑 Configuração da Chave da API (Gemini AI)

1. Abra o arquivo `.env` no Bloco de Notas ou editor de sua preferência.
2. Insira sua chave da API:
   ```env
   GEMINI_API_KEY=sua_chave_aqui
   ```
3. Salve o arquivo. O sistema carregará a chave automaticamente na próxima inicialização.

---

## 📁 Estrutura dos Arquivos Principais

| Arquivo / Pasta | Finalidade |
| :--- | :--- |
| **`INICIAR_SISTEMA.bat`** | **Executável principal de 1 clique** para abrir o sistema. |
| **`EXECUTAR.bat`** | Atalho rápido para o iniciador. |
| **`INSTALAR_DEPENDENCIAS.bat`** | Instalador e atualizador de pacotes em 1 clique. |
| **`app.py`** | Código-fonte unificado da aplicação desktop (PySide6 / Qt). |
| **`requirements.txt`** | Lista oficial de dependências do Python. |
| **`.env`** | Configuração local da chave da API (privado, não comitado). |
| **`.env.example`** | Modelo de referência das variáveis de ambiente. |
| **`prints_cotacoes/`** | Pasta onde são salvos os comprovantes com carimbo oficial. |
| **`Modelo_Itens_Cotacao.xlsx`** | Modelo de planilha de itens para importação de termos de referência. |

---

## 🏛️ Funcionalidades Integradas

- **Pesquisa Comercial Automatizada:** Coleta de cotações em lojas e fornecedores com Playwright.
- **Auditoria por IA (Gemini 3.1 Flash Lite):** Conferência de especificações técnicas, equivalência e conversão de embalagens.
- **Validação de Fornecedores e CNPJ:** Consulta automática na base oficial da Receita Federal.
- **Carimbo Oficial Monocromático:** Certificação gráfica das páginas capturadas com data, hora, IP/origem e hash.
- **Exportação Multiformato:** Planilha consolidada em Excel (.xlsx) e pacote digital completo (.zip).
