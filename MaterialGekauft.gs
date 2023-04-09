// Stand 09.04.23
const MATERIAL_GEKAUFT_START_ROW = 7s;

// Es werden Daten nur unten angehängt, bestehende Daten werden nicht verändert/ersetzt!
function fillAndAddMaterialGekauft() {
    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetGekauft = SpreadsheetApp.getActive().getSheetByName('Material gekauft');

    //  Sheet Gekauft: bisher ausgefüllte Gegenstände und max Zeilen anzahl ermitteln
    var uebertrageneGegenstaende = {};
    var indexFuerNeueDaten = MATERIAL_GEKAUFT_START_ROW;

    var rangeMaterialDataVorhanden = sheetGekauft.getRange(MATERIAL_GEKAUFT_START_ROW, 2, MAX_ROWS, 1).getValues();
    rangeMaterialDataVorhanden.forEach(function (row) {
        let gegenstandName = row[0];
        if (gegenstandName) {
            uebertrageneGegenstaende[gegenstandName] = "yes";
            indexFuerNeueDaten++;
        }
    });

    //  Zu schreibende Gegenstände aus Gesamtsheet ermitteln
    var gegenstandNameZuAnzahlZuKaufen = {};

    var rangeGesamtData = sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 2, MAX_ROWS, 12).getValues();
    rangeGesamtData.forEach(function (row) {
        let gegenstandName = row[0];
        let zuKaufen = row[4];
        if (zuKaufen == 'x') {
            let anzahlBenoetigt = row[9];
            if (anzahlBenoetigt == 0) {
                console.log('Anzahl benötigt ist 0 für Gegenstand ' + gegenstandName);
                return;
            }
            let anzahlVonResortGestellt = row[11];
            let anzahlZuKaufen = anzahlBenoetigt - anzahlVonResortGestellt;
            if (anzahlZuKaufen > 0) {
                gegenstandNameZuAnzahlZuKaufen[gegenstandName] = anzahlZuKaufen;
            } else {
                console.log("Nichts zu kaufen für Gegenstand " + gegenstandName);
            }
        }
    });

    //  Gegenstände die noch nicht da sind unter bestehenden Daten einfügen
    var gegenstandNameInsert = [];
    var anzahlZuKaufenInsert = [];
    for (const [gegenstandName, anzahlZuKaufen] of Object.entries(gegenstandNameZuAnzahlZuKaufen)) {
        let schonVorhanden = uebertrageneGegenstaende[gegenstandName];
        if (!schonVorhanden) {
            gegenstandNameInsert.push(gegenstandName);
            anzahlZuKaufenInsert.push(anzahlZuKaufen);
        }
    }

    let anzahlZeilenToInsert = gegenstandNameInsert.length;

    if (gegenstandNameInsert.length > 0) {
        console.log("Schreibe ", anzahlZeilenToInsert, " Zeilen in Gekauft Sheet ab Index ", indexFuerNeueDaten);
        sheetGekauft.getRange(indexFuerNeueDaten, 2, anzahlZeilenToInsert, 1).setValues(convertIn2dArray(gegenstandNameInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 4, anzahlZeilenToInsert, 1).setValues(convertIn2dArray(anzahlZuKaufenInsert));

    }

    // Stand befüllen
    printStandInZelle("B2", sheetGekauft);
    Browser.msgBox(anzahlZeilenToInsert + " Datensätze eingefügt/ergänzt");

}
