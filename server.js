// server.js

const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const { google } = require('googleapis');

const app = express();
// Armazena o PDF em memória para processar
const upload = multer({ storage: multer.memoryStorage() });

// ======================
// 1. CONFIGURAÇÃO SHEETS
// ======================
const SPREADSHEET_ID = 'SEU_SPREADSHEET_ID_AQUI';
const SHEET_NAME     = 'Sheet1';  // ajuste para o nome da sua aba

// Autenticação via conta de serviço (credentials.json na raiz do projeto)
const auth = new google.auth.GoogleAuth({
  keyFile: 'credentials.json',
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });

// ======================
// 2. ROTA /extract
// ======================
app.post('/extract', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ sucesso: false, erro: 'Nenhum arquivo enviado' });
    }

    // 2.1. Extrai texto do PDF
    const data = await pdfParse(req.file.buffer);
    const text = data.text;

    // 2.2. Extração de campos via regex (ajuste as expressões conforme seu layout)
    const numeroNota = (text.match(/Nota Fiscal[:\s]+(\d+)/i) || [])[1] || '';
    // Exemplo de outro campo:
    // const dataEmissao = (text.match(/Data Emissão[:\s]+(\d{2}\/\d{2}\/\d{4})/i) || [])[1] || '';

    // 2.3. Prepara os valores para inserir na planilha
    const values = [
      [
        numeroNota,
        // dataEmissao,
        // ...outros campos
      ]
    ];

    // 2.4. Insere nova linha na planilha
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      resource: { values }
    });

    // 2.5. Retorna JSON ao front-end
    res.json({
      sucesso: true,
      dados: { numeroNota /*, dataEmissao */ }
    });

  } catch (err) {
    console.error('Erro no /extract:', err);
    res.status(500).json({ sucesso: false, erro: err.message });
  }
});

// ======================
// 3. INICIA SERVIDOR
// ======================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Servidor rodando em http://localhost:${PORT}`);
});