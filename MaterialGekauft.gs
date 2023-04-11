// Stand 09.04.23
const MATERIAL_GEKAUFT_START_ROW = 7;

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

    //  Gegenstände, die noch nicht in "Material Gekauft" Liste enthalten sind ermitteln
    var gegenstandNameInsert = [];
    var anzahlZuKaufenInsert = [];
    for (const[gegenstandName, anzahlZuKaufen]of Object.entries(gegenstandNameZuAnzahlZuKaufen)) {
        let schonVorhanden = uebertrageneGegenstaende[gegenstandName];
        if (!schonVorhanden) {
            gegenstandNameInsert.push(gegenstandName);
            anzahlZuKaufenInsert.push(anzahlZuKaufen);
        }
    }

    // Link für Kauf + Transport aus Resortliste Gesamt von relevanten Gegenstaenden ermitteln
    let anzahlZeilenToInsert = gegenstandNameInsert.length;
    if (anzahlZeilenToInsert > 0) {
        var gegenstandNameZuLink = {};
        var gegenstandNameZuTransport = {};
        var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

        var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(RESORT_GESAMT_LISTE_START_ROW, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 11).getValues();
        console.log(rangeGesamtData.length);
        var resortListeGesamtFiltered = rangeResortlisteGesamt.filter(function (row) {
            return gegenstandNameInsert.includes(row[0]);
        });

        resortListeGesamtFiltered.forEach(function (row) {
            mergeMap(gegenstandNameZuLink, row[0], row[9]);
            mergeMap(gegenstandNameZuTransport, row[0], row[10]);
        });

        // Werte in gleiche Reihenfolge bringen wie einzufügenden Gegenstände
        var linkInsert = [];
        var transportInsert = [];
        gegenstandNameInsert.forEach(gegenstand => {
            let link = gegenstandNameZuLink[gegenstand];
            if (link) {
                linkInsert.push(link);
            } else {
                linkInsert.push("");
            }

            let transport = gegenstandNameZuTransport[gegenstand];
            if (transport) {
                transportInsert.push(transport);
            } else {
                transportInsert.push("");
            }
        });

        // Daten schreiben
        console.log("Schreibe ", anzahlZeilenToInsert, " Zeilen in Gekauft Sheet ab Index ", indexFuerNeueDaten);
        sheetGekauft.getRange(indexFuerNeueDaten, 2, anzahlZeilenToInsert, 1).setValues(convertIn2dArray(gegenstandNameInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 4, anzahlZeilenToInsert, 1).setValues(convertIn2dArray(anzahlZuKaufenInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 10, anzahlZeilenToInsert, 1).setValues(convertIn2dArrayAndJoinData(linkInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 13, anzahlZeilenToInsert, 1).setValues(convertIn2dArrayAndJoinData(transportInsert));
    }

    // Stand befüllen
    printStandInZelle("B2", sheetGekauft);
    Browser.msgBox(anzahlZeilenToInsert + " Datensätze eingefügt/ergänzt");

}
