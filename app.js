const express = require('express');
const excel = require('excel4node');

const app = express();

let data = require('./data');

    app.get('/', (req, res) => {
    res.end('Server is up');
});

app.get('/report', (req, res) => {
    let wb = new excel.Workbook();
    let ws_total = wb.addWorksheet('Общий');
    let styles = require('./styles')(wb);
    console.log(data.length);
    let start_row = 2;

    ws_total.cell(1, 1, 1, 11, true)
        .string('This sheet is provided by PPPoker Co., Ltd. and is based on the derivative data of the virtual game currency , which is only a reference and does not have legal effect.')
        .style(styles.common);

    data.forEach((table, index) => {
        console.log(index, table);
        ws_total.cell(start_row, 1, (start_row + table.players.length + 5), 1, true)
            .string(table.date)
            .style(styles.header_top_left);
        ws_total.cell(start_row, 2, start_row, 15, true)
            .string(`Start time: ${table.data.start_time} By ${table.data.owner.name}(${table.data.owner.player_id})`)
            .style(styles.header_top);
        ws_total.cell(start_row + 1, 2, start_row + 1, 15, true)
            .string(`Table name: ${table.data.name}`)
            .style(styles.header);
        ws_total.cell(start_row + 2, 2, start_row + 2, 15, true)
            .string(`Table information: ${table.data.info}`)
            .style(styles.header);
        ws_total.cell(start_row + 3, 2, start_row + 4, 2, true)
            .string('ID игрока')
            .style(styles.title_green);
        ws_total.cell(start_row + 3, 3, start_row + 4, 3, true)
            .string('Ник')
            .style(styles.title_green);
        ws_total.cell(start_row + 3, 4, start_row + 4, 4, true)
            .string('Игровое имя')
            .style(styles.title_green);
        ws_total.cell(start_row + 3, 5, start_row + 4, 5, true)
            .string('Бай-ин с Q-фишками')
            .style(styles.title_green);
        ws_total.cell(start_row + 3, 6, start_row + 4, 6, true)
            .string('Раздачи')
            .style(styles.title_green);
        ws_total.cell(start_row + 3, 7, start_row + 3, 10, true)
            .string('Выигрыш игрока')
            .style(styles.title_red);
        ws_total.cell(start_row + 4, 7)
            .string('Общий')
            .style(styles.title_red);
        ws_total.cell(start_row + 4, 8)
            .string('От соперников')
            .style(styles.title_red);
        ws_total.cell(start_row + 4, 9)
            .string('От джекпота')
            .style(styles.title_red);
        ws_total.cell(start_row + 4, 10)
            .string('От страховки')
            .style(styles.title_red);
        ws_total.cell(start_row + 3, 11, start_row + 3, 15, true)
            .string('Доход клуба')
            .style(styles.title_green);
        ws_total.cell(start_row + 4, 11)
            .string('Общий')
            .style(styles.title_green);
        ws_total.cell(start_row + 4, 12)
            .string('Комиссия')
            .style(styles.title_green);
        ws_total.cell(start_row + 4, 13)
            .string('Комиссия джекпота')
            .style(styles.title_green);
        ws_total.cell(start_row + 4, 14)
            .string('Выплаты джекпота')
            .style(styles.title_green);
        ws_total.cell(start_row + 4, 15)
            .string('Страхование')
            .style(styles.title_green);
        ws_total.cell(start_row + 3, 16)
            .string('')
            .style(styles.right_end);
        ws_total.cell(start_row + 4, 16)
            .string('')
            .style(styles.right_end);
        table.players.forEach((player, index) => {
            ws_total.cell(start_row + 5 + index, 2)
                .number(player.id)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 3)
                .string(player.nick)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 4)
                .string(player.name)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 5)
                .number(player.buy_in)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 6)
                .number(player.circle)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 7)
                .number(player.player_win.total)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 8)
                .number(player.player_win.enemy)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 9)
                .number(player.player_win.jackpot)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 10)
                .number(player.player_win.insurance)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 11)
                .number(player.club_win.total)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 12)
                .number(player.club_win.tax)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 13)
                .number(player.club_win.jackpot_tax)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 14)
                .number(player.club_win.jackpot_pay)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 15)
                .number(player.club_win.insurance)
                .style(styles.common);
            ws_total.cell(start_row + 5 + index, 16)
                .string('')
                .style(styles.right_end);
        });

        ws_total.cell(start_row + table.players.length + 5, 2, start_row + table.players.length + 5, 5, true)
            .style(styles.border_bottom);

        ws_total.cell(start_row + table.players.length + 5, 6)
            .string('Итог')
            .style(styles.border_bottom);

        ws_total.cell(start_row + table.players.length + 5, 7)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 7) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 7) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 8)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 8) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 8) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 9)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 9) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 9) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 10)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 10) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 10) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 11)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 11) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 11) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 12)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 12) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 12) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 13)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 13) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 13) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 14)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 14) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 14) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 15)
            .formula('SUM(' + excel.getExcelCellRef(start_row + 5, 15) + ':' + excel.getExcelCellRef(start_row + 4 + table.players.length, 15) + ')')
            .style(styles.border_bottom);
        ws_total.cell(start_row + table.players.length + 5, 16)
            .string('')
            .style(styles.right_end);


        start_row = start_row + (7 + table.players.length);
    });




    /*ws_total.cell(2, 1, 9, 1, true)
        .date(new Date('2019/10/04UTC +0300'))
        .style(styles.header_top_left);
    ws_total.cell(2, 2, 2, 15, true)
        .string('Start Time: 17:41 By iShimon(1123390)')
        .style(styles.header_top);
    ws_total.cell(3, 2, 3, 15, true)
        .string('Table Name: Безымянный')
        .style(styles.header);
    ws_total.cell(4, 2, 4, 15, true)
        .string('Table Information: 0.5/1  NLH  5%  1.0BB  12.0h')
        .style(styles.header);*/

    wb.write('Ex.xlsx', res);
});

app.listen(5000, () => {
    console.log('Server is up!');
});