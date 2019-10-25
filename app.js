const express = require('express');
const excel = require('excel4node');

const app = express();

let data = require('./data');

    app.get('/', (req, res) => {
    res.end('Server is up');
});

app.get('/report', (req, res) => {
    let wb = new excel.Workbook({
        dateFormat: 'm/d/yy hh:mm:ss',
    });
    let styles = require('./styles')(wb);
    require('./sheets/game')(excel, wb, data.games, styles.games);
    require('./sheets/transactions')(excel, wb, data.transactions, styles.transactions);

    wb.write('report.xlsx', res);
});

app.listen(5000, () => {
    console.log('Server is up!');
});