module.exports= (excel, wb, data, styles) => {
    let wb_sheet = wb.addWorksheet('История операций', {
        sheetFormat: {
            defaultColWidth: 14,
            defaultRowHeight: 15,
        }
    });
    let start_row_for_game = 2;

    wb_sheet.cell(1, 1, 1, 11, true)
        .string('This sheet is provided by QPoker and is based on the derivative data of the virtual game currency , which is only a reference and does not have legal effect.')
        .style(styles.common);

    wb_sheet.cell(2, 1, 3, 1, true)
        .string('Время')
        .style(styles.header);
    wb_sheet.cell(2, 2, 2, 4, true)
        .string('Отправитель')
        .style(styles.header);
    wb_sheet.cell(2, 5, 2, 7, true)
        .string('Получатель')
        .style(styles.header);
    wb_sheet.cell(2, 8, 2, 9, true)
        .string('Дать кредит')
        .style(styles.header);
    wb_sheet.cell(2, 10, 2, 11, true)
        .string('Фишки')
        .style(styles.header);
    wb_sheet.cell(3, 2)
        .string('ID игрока')
        .style(styles.header_green);
    wb_sheet.cell(3, 3)
        .string('Ник')
        .style(styles.header_green);
    wb_sheet.cell(3, 4)
        .string('Игровое имя')
        .style(styles.header_green);
    wb_sheet.cell(3, 5)
        .string('ID игрока')
        .style(styles.header_green);
    wb_sheet.cell(3, 6)
        .string('Ник')
        .style(styles.header_green);
    wb_sheet.cell(3, 7)
        .string('Игровое имя')
        .style(styles.header_green);
    wb_sheet.cell(3, 8)
        .string('Выдать')
        .style(styles.header_green);
    wb_sheet.cell(3, 9)
        .string('Вернуть')
        .style(styles.header_green);
    wb_sheet.cell(3, 10)
        .string('Выдать')
        .style(styles.header_green);
    wb_sheet.cell(3, 11)
        .string('Вернуть')
        .style(styles.header_green);

    data.forEach((transaction, index) => {
        wb_sheet.cell(4 + index, 1)
            .string(transaction.date)
            .style(styles.row);
        wb_sheet.cell(4 + index, 2)
            .number(transaction.sender.player_id)
            .style(styles.row);
        wb_sheet.cell(4 + index, 3)
            .string(transaction.sender.player_name)
            .style(styles.row);
        wb_sheet.cell(4 + index, 4)
            .string(transaction.sender.club_member_name)
            .style(styles.row);
        wb_sheet.cell(4 + index, 5)
            .number(transaction.recipient.player_id)
            .style(styles.row);
        wb_sheet.cell(4 + index, 6)
            .string(transaction.recipient.player_name)
            .style(styles.row);
        wb_sheet.cell(4 + index, 7)
            .string(transaction.recipient.club_member_name)
            .style(styles.row);
        wb_sheet.cell(4 + index, 8)
            .number(transaction.credit.give)
            .style(styles.row);
        wb_sheet.cell(4 + index, 9)
            .number(transaction.credit.back)
            .style(styles.row);
        wb_sheet.cell(4 + index, 10)
            .number(transaction.chips.give)
            .style(styles.row);
        wb_sheet.cell(4 + index, 11)
            .number(transaction.chips.back)
            .style(styles.row);
    });
};