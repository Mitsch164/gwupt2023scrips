// Stand 08.04.23

const MAX_ROWS = 500;
const MAX_ROWS_RESORTLISTE_KOMPLETT = 2000;
const RESORT_LISTE_START_ROW = 8;
const RESORT_GESAMT_LISTE_START_ROW = 18;
const GESAMT_LISTE_START_ROW = 7;

function printStandInZelle(zelleName, sheet) {
    var standFormatted = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.YYYY");
    var standCell = sheet.getRange(zelleName).getCell(1, 1);
    standCell.setValue(standFormatted);
}

function mergeMap(map, key, valueData) {
    let valueInMap = map[key];
    if (!valueInMap) {
        valueInMap = [];
    }
    valueInMap.push(valueData);
    map[key] = valueInMap;
}
