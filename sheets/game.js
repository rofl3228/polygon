module.exports = (excel, wb, data, styles) => {
    let wb_sheet = wb.addWorksheet('Игры', {
        sheetFormat: {
            defaultColWidth: 14,
            defaultRowHeight: 17,
        }
    });
    let start_row_for_game = 2;

    wb_sheet.cell(1, 1, 1, 11, true)
        .string('This sheet is provided by QPoker and is based on the derivative data of the virtual game currency , which is only a reference and does not have legal effect.')
        .style(styles.common);

    data.forEach((table) => {
        //Table headder
        wb_sheet.cell(start_row_for_game, 1, (start_row_for_game + table.participants.length + 5), 1, true)
            .date(new Date(table.game.start_time.substring(0, 10)))
            .style(styles.header_top_left);
        wb_sheet.cell(start_row_for_game, 2, start_row_for_game, 15, true)
            .string(`Start time: ${table.game.start_time.substring(11,16)} By ${table.player.name}(${table.player.id})`)
            .style(styles.header_top);
        wb_sheet.cell(start_row_for_game + 1, 2, start_row_for_game + 1, 15, true)
            .string(`Table name: -`)
            .style(styles.header);
        wb_sheet.cell(start_row_for_game + 2, 2, start_row_for_game + 2, 15, true)
            .string(`Table information: ${table.game.blinds.small}/${table.game.blinds.big} ${table.game.game_type} ${table.game.rake.fee}% ${table.game.rake.cap}BB ${table.game.duration}h`)
            .style(styles.header);
        //Column headers
        wb_sheet.cell(start_row_for_game + 3, 2, start_row_for_game + 4, 2, true)
            .string('ID игрока')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 3, 3, start_row_for_game + 4, 3, true)
            .string('Ник')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 3, 4, start_row_for_game + 4, 4, true)
            .string('Игровое имя')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 3, 5, start_row_for_game + 4, 5, true)
            .string('Бай-ин с Q-фишками')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 3, 6, start_row_for_game + 4, 6, true)
            .string('Раздачи')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 3, 7, start_row_for_game + 3, 10, true)
            .string('Выигрыш игрока')
            .style(styles.title_red);
        wb_sheet.cell(start_row_for_game + 4, 7)
            .string('Общий')
            .style(styles.title_red);
        wb_sheet.cell(start_row_for_game + 4, 8)
            .string('От соперников')
            .style(styles.title_red);
        wb_sheet.cell(start_row_for_game + 4, 9)
            .string('От джекпота')
            .style(styles.title_red);
        wb_sheet.cell(start_row_for_game + 4, 10)
            .string('От страховки')
            .style(styles.title_red);
        wb_sheet.cell(start_row_for_game + 3, 11, start_row_for_game + 3, 15, true)
            .string('Доход клуба')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 4, 11)
            .string('Общий')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 4, 12)
            .string('Комиссия')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 4, 13)
            .string('Комиссия джекпота')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 4, 14)
            .string('Выплаты джекпота')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 4, 15)
            .string('Страхование')
            .style(styles.title_green);
        wb_sheet.cell(start_row_for_game + 3, 16)
            .string('')
            .style(styles.right_end);
        wb_sheet.cell(start_row_for_game + 4, 16)
            .string('')
            .style(styles.right_end);
        //Players data
        table.participants.forEach((participant, index) => {
            wb_sheet.cell(start_row_for_game + 5 + index, 2)
                .number(participant.player.player_id)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 3)
                .string(participant.player.club_member_name)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 4)
                .string(participant.player.player_name)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 5)
                .number(participant.buy_in_sum)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 6)
                .number(participant.deals_amount)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 7)
                .number(participant.win_sum)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 8)
                .number(0/*player.player_win.enemy*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 9)
                .number(0/*player.player_win.jackpot*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 10)
                .number(0/*player.player_win.insurance*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 11)
                .number(0/*player.club_win.total*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 12)
                .number(0/*player.club_win.tax*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 13)
                .number(0/*player.club_win.jackpot_tax*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 14)
                .number(0/*player.club_win.jackpot_pay*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 15)
                .number(0/*player.club_win.insurance*/)
                .style(styles.common);
            wb_sheet.cell(start_row_for_game + 5 + index, 16)
                .string('')
                .style(styles.right_end);
        });

        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 2, start_row_for_game + table.participants.length + 5, 5, true)
            .style(styles.border_bottom);

        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 6)
            .string('Итог')
            .style(styles.border_bottom);

        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 7)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 7) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 7) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 8)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 8) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 8) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 9)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 9) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 9) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 10)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 10) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 10) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 11)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 11) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 11) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 12)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 12) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 12) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 13)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 13) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 13) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 14)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 14) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 14) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 15)
            .formula('SUM(' + excel.getExcelCellRef(start_row_for_game + 5, 15) + ':' + excel.getExcelCellRef(start_row_for_game + 4 + table.participants.length, 15) + ')')
            .style(styles.border_bottom);
        wb_sheet.cell(start_row_for_game + table.participants.length + 5, 16)
            .string('')
            .style(styles.right_end);


        start_row_for_game = start_row_for_game + (7 + table.participants.length);
    });
    return wb_sheet;
};