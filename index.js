const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const cors = require('cors');
const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname)));
app.use(bodyParser.json());
app.use(cors())
app.post('/update', (req, res) => {
    const newData = req.body;
    const workbook = XLSX.readFile('emails.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const oldData = XLSX.utils.sheet_to_json(worksheet);

    const oldDataMap = {};
    oldData.forEach((row, index) => {
        oldDataMap[row.id] = { ...row, row: index + 1 };
    });


    const newDataMap = {};
    newData.forEach((row) => {
        newDataMap[row.id] = row;
    });


    for(let id in oldDataMap){
        if (newDataMap[id]) {
        
            const oldRow = oldDataMap[id];
            XLSX.utils.sheet_add_aoa(worksheet, [[newDataMap[id]]], {origin: `A${oldRow.row}`});
            delete newDataMap[id];
        }
    }

    XLSX.writeFile(workbook, 'emails.xlsx');

    res.send('Archivo de Excel sincronizado con éxito!');
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});
app.get('/clientessoloio', (req, res)=>{
    res.sendFile(path.join(__dirname, 'admin.html'));
})
app.get('/preview', (req, res) => {
    const workbook = XLSX.readFile('emails.xlsx');
    const sheet_name_list = workbook.SheetNames;
    const html_string = XLSX.utils.sheet_to_html(workbook.Sheets[sheet_name_list[0]]);
    res.send(html_string);
});
app.get('/preview2', (req, res) => {
    const workbook = XLSX.readFile('emails.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    res.json(jsonData);
});

app.get('/download', (req, res) => {
    res.download('emails.xlsx');
});
app.post('/submit-email', (req, res) => {
    const email = req.body.email;
    const workbook = XLSX.readFile('emails.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    let nextRow;
    if (worksheet['!ref']) {
        nextRow = XLSX.utils.decode_range(worksheet['!ref']).e.r + 2;
    } else {
        nextRow = 1;
    }
    XLSX.utils.sheet_add_aoa(worksheet, [[email]], {origin: `A${nextRow}`});
    XLSX.writeFile(workbook, 'emails.xlsx');
});
app.listen(3000, () => console.log('Servidor corriendo en http://localhost:3000'));
