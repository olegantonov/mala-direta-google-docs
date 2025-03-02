# Mala Direta Google Docs

Este projeto automatiza a criação de mala direta para impressão de envelopes a partir de dados armazenados em uma planilha do **Google Sheets**. Ele gera um **Google Docs** formatado, preenchendo os dados automaticamente e atualizando a planilha com o status do processamento.

## 📌 Funcionalidades
- **Busca dados do Google Sheets** para preencher os destinatários.
- **Cria um Google Docs** formatado com os endereços para impressão.
- **Atualiza automaticamente o status** de cada entrada processada.
- **Exibe um link para o documento gerado** ao final da execução.
- **Adiciona um menu personalizado** ao Google Sheets para facilitar a execução.

## 🛠️ Tecnologias Utilizadas
- **Google Apps Script**
- **Google Sheets API**
- **Google Docs API**
- **Google Drive API**

## 🚀 Como Usar
1. **Crie um Google Sheets** e configure os dados nas colunas correspondentes:
   - **Coluna M (12)** → Status do processamento.
   - **Coluna P (15)** → Nome do destinatário.
   - **Coluna Y (24)** → Endereço.
   - **Coluna Z (25)** → CEP.
   - **Coluna AB (27)** → Cidade.
   - **Coluna F (5)** → Instituição.
   - **Coluna R (17)** → Pronome de tratamento.
   - **Coluna O (14)** → Sexo.
   - **Coluna 13** → Status de atualização.

2. **Configure o modelo do Google Docs**:
   - Defina o **ID do modelo** de documento na constante:
     ```js
     var modeloId = "SEU_ID_AQUI";
     ```

3. **Adicione o código no Google Apps Script** do Google Sheets:
   - Acesse `Extensões` > `Apps Script` e cole o código no editor.

4. **Execute a função `onOpen()`** para adicionar o menu "Exportar" na planilha.

5. **Clique no menu "Exportar"** e selecione **"IMPRESSÃO ENVELOPE DL"** para gerar os documentos.

## 📄 Licença
Este projeto é distribuído sob a licença MIT.

---

### ✍️ Autor
**Daniel Marques**
