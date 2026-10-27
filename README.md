# Sistema de Pesquisa de Preços de Mercado — Pregão SRP

Sistema corporativo para automação de pesquisa de preços de mercado, auditoria por inteligência artificial e geração de comprovantes oficiais com carimbo monocromático de autenticidade, em estrita conformidade com a **Lei Federal nº 14.133/2021** (Nova Lei de Licitações) e Instrução Normativa SEGES/ME.

Disponível em dois modos de uso: **Aplicativo Desktop (PySide6 / Qt)** e **Servidor Web Local (Navegador Web)**.

---

## 🚀 Como Executar em 1 Clique (Simples e Direto)

Basta dar um **duplo clique** em qualquer um dos arquivos abaixo na pasta do projeto:

### Opção 1: Usar no Navegador Web (Recomendado)
- **`INICIAR_SERVIDOR_WEB.bat`**  
  *Inicia o servidor web local e abre automaticamente o sistema no Google Chrome, Microsoft Edge ou seu navegador padrão em `http://127.0.0.1:5000/`.*

### Opção 2: Usar como Aplicativo Desktop
- **`INICIAR_SISTEMA.bat`** ou **`EXECUTAR.bat`**  
  *Abre a interface gráfica nativa para desktop (Windows Qt).*

Ambos os iniciadores:
1. Detectam o Python automaticamente no seu Windows.
2. Criam o arquivo `.env` de configuração se ainda não existir.
3. Instalam as dependências necessárias automaticamente na primeira execução.
4. Iniciam o sistema sem que você precise digitar comandos no terminal.

---

## 🛠️ Execução por Linha de Comando (Opcional)

Se preferir utilizar o terminal:

```bash
# Iniciar a versão Web no navegador:
python app.py --web
# ou:
python servidor_web.py

# Iniciar a versão Desktop (Qt):
python app.py
```

---

## 🔑 Configuração da Chave da API (Gemini AI)

1. Abra o arquivo `.env` no Bloco de Notas ou editor de sua preferência.
2. Insira sua chave da API:
   ```env
   GEMINI_API_KEY=sua_chave_aqui
   ```
3. Salve o arquivo. O sistema carregará a chave automaticamente na inicialização.

---

## 📁 Estrutura dos Arquivos Principais

| Arquivo / Pasta | Finalidade |
| :--- | :--- |
| **`INICIAR_SERVIDOR_WEB.bat`** | **Executável de 1 clique** para abrir a versão Web no navegador. |
| **`INICIAR_SISTEMA.bat`** | **Executável de 1 clique** para abrir a versão Desktop (Qt). |
| **`EXECUTAR.bat`** | Atalho rápido para o aplicativo desktop. |
| **`INSTALAR_DEPENDENCIAS.bat`** | Instalador e reparador de dependências em 1 clique. |
| **`servidor_web.py`** | Servidor web local (Flask) e API REST de alta performance. |
| **`templates/index.html`** | Interface Web responsiva com o mesmo layout corporativo do desktop. |
| **`app.py`** | Código-fonte da aplicação desktop e motor de pesquisa/IA compartilhado. |
| **`requirements.txt`** | Lista oficial de dependências do Python. |
| **`.env`** | Configuração local da chave da API (privado, não comitado). |
| **`.env.example`** | Modelo de referência das variáveis de ambiente. |
| **`prints_cotacoes/`** | Pasta onde são salvos os comprovantes com carimbo oficial. |
| **`Modelo_Itens_Cotacao.xlsx`** | Modelo de planilha de itens para importação de termos de referência. |

---

## 🏛️ Funcionalidades Integradas

- **Layout Idêntico e Responsivo:** Mesma hierarquia visual, cores institucionais, KPIs compactos, tabela expansível e painel de detalhes tanto na versão Web quanto no Desktop.
- **Pesquisa Comercial Automatizada:** Coleta de cotações em lojas e fornecedores com Playwright.
- **Auditoria por IA (Gemini 3.1 Flash Lite):** Conferência de especificações técnicas, equivalência e conversão de embalagens.
- **Validação de Fornecedores e CNPJ:** Consulta automática na base oficial da Receita Federal.
- **Visualizador de Comprovantes:** Zoom dinâmico, ajuste de largura/janela e navegação entre lotes.
- **Carimbo Oficial Monocromático:** Certificação gráfica das páginas capturadas com data, hora, IP/origem e hash.
- **Exportação Multiformato:** Planilha consolidada em Excel (.xlsx) e pacote digital completo (.zip).
