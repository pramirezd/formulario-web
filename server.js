const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

app.use(bodyParser.json());
app.use(express.static('public'));

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/public/main.html');
});

app.post('/submit', (req, res) => {
    const formData = req.body;

    fs.readFile('data.json', (err, data) => {
        if (err) throw err;
        const jsonData = JSON.parse(data);
        jsonData.push(formData);

        fs.writeFile('data.json', JSON.stringify(jsonData, null, 2), (err) => {
            if (err) throw err;
            res.json({ message: 'Datos enviados con éxito' });
        });
    });
});

app.get('/download-report', (req, res) => {
    fs.readFile('data.json', async (err, data) => {
        if (err) throw err;
        const jsonData = JSON.parse(data);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Reportes');

        worksheet.columns = [
            { header: 'Nombre', key: 'name', width: 20 },
            { header: 'Edad', key: 'age', width: 10 },
            { header: 'Género', key: 'gender', width: 15 },
            { header: 'Aficiones', key: 'hobbies', width: 30 },
        ];

        jsonData.forEach((entry) => {
            worksheet.addRow({
                name: entry.name,
                age: entry.age,
                gender: entry.gender,
                hobbies: entry.hobbies.join(', ')
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=data.xlsx');

        await workbook.xlsx.write(res);
        res.end();
    });
});

app.listen(PORT, () => {
    console.log(`Servidor ejecutándose en http://localhost:${PORT}`);
});
