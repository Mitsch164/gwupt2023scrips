// Stand 24.03.23

const MAX_ROWS = 500;

function printStandInZelle(zelleName, sheet){
    var standFormatted = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.YYYY");
    var standCell = sheet.getRange(zelleName).getCell(1,1);
    standCell.setValue(standFormatted);
}