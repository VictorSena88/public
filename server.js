const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');

const app = express();
const port = 3000;

// Middleware para tratar dados JSON
app.use(bodyParser.json());
app.use(express.static('public')); // Para servir arquivos estáticos (HTML, CSS, JS)

app.post('/gerar-excel', (req, res) => {
    const alunos = req.body.alunos;

    // Criação da planilha Excel
    const ws = xlsx.utils.json_to_sheet(alunos);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Alunos e Notas');

    // Cria um buffer do arquivo Excel
    const buffer = xlsx.write(wb, { bookType: 'xlsx', type: 'buffer' });

    // Envia o arquivo Excel como resposta
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=alunos_notas.xlsx');
    res.send(buffer);
});

app.listen(port, () => {
    console.log(`Servidor rodando em http://localhost:${port}`);
});