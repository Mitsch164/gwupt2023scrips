// Stand 16.04.23

// TODO Spalte Resort eingefügt, aber keine Ausführung mehr notwendig

const MATERIAL_GEKAUFT_START_ROW = 7;

// Es werden Daten nur unten angehängt, bestehende Daten werden nicht verändert/ersetzt!
function fillAndAddMaterialGekauft() {
    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetGekauft = SpreadsheetApp.getActive().getSheetByName('Material gekauft');

    //  Sheet Gekauft: bisher ausgefüllte Gegenstände und max Zeilen anzahl ermitteln
    var uebertrageneGegenstaende = [];
    var indexFuerNeueDaten = MATERIAL_GEKAUFT_START_ROW;

    var rangeMaterialDataVorhanden = sheetGekauft.getRange(MATERIAL_GEKAUFT_START_ROW, 2, MAX_ROWS, 1).getValues();
    rangeMaterialDataVorhanden.forEach(function (row) {
        let gegenstandName = row[0];
        if (gegenstandName) {
            uebertrageneGegenstaende.push(gegenstandName);
            indexFuerNeueDaten++;
        }
    });

    //  Zu schreibende Gegenstände aus Gesamtsheet ermitteln
    var gegenstandNameZuUebertragen = [];

    var rangeGesamtData = sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 2, MAX_ROWS, 13).getValues();
    rangeGesamtData.forEach(function (row) {
        let gegenstandName = row[0];
        let zuKaufen = row[4];
        if (zuKaufen == 'x') {
            let anzahlBenoetigt = row[10];
            gegenstandNameZuUebertragen.push(gegenstandName);
        }
    });

    // Gegenstände, die noch nicht in "Material Gekauft" Liste enthalten sind ermitteln
    var gegenstandNameInsert = [];
    gegenstandNameInsert = gegenstandNameZuUebertragen.filter(gegenstand => !uebertrageneGegenstaende.includes(gegenstand));
    console.log(gegenstandNameInsert);

    // Link für Kauf + Transport + Transport Besonderheiten aus Resortliste Gesamt von relevanten Gegenstaenden ermitteln
    let anzahlZeilenToInsert = gegenstandNameInsert.length;
    if (anzahlZeilenToInsert > 0) {
        var gegenstandNameZuLink = {};
        var gegenstandNameZuTransport = {};
        var gegenstandNameZuBesonderheitTransport = {};

        var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

        var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(RESORT_GESAMT_LISTE_START_ROW, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 12).getValues();
        var resortListeGesamtFiltered = rangeResortlisteGesamt.filter(function (row) {
            return gegenstandNameInsert.includes(row[0]);
        });

        resortListeGesamtFiltered.forEach(function (row) {
            mergeMap(gegenstandNameZuLink, row[0], row[9]);
            mergeMap(gegenstandNameZuTransport, row[0], row[10]);
            mergeMap(gegenstandNameZuBesonderheitTransport, row[0], row[11]);
        });

        // Werte in gleiche Reihenfolge bringen wie einzufügenden Gegenstände
        var linkInsert = [];
        var transportInsert = [];
        var transportBesonderheitInsert = [];
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

            let transportBesonderheit = gegenstandNameZuBesonderheitTransport[gegenstand];
            if (transportBesonderheit) {
                transportBesonderheitInsert.push(transportBesonderheit);
            } else {
                transportBesonderheitInsert.push("");
            }
        });

        // Daten schreiben
        console.log("Schreibe ", anzahlZeilenToInsert, " Zeilen in Gekauft Sheet ab Index ", indexFuerNeueDaten);
        sheetGekauft.getRange(indexFuerNeueDaten, 2, anzahlZeilenToInsert, 1).setValues(convertIn2dArray(gegenstandNameInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 12, anzahlZeilenToInsert, 1).setValues(convertIn2dArrayAndJoinData(linkInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 16, anzahlZeilenToInsert, 1).setValues(convertIn2dArrayAndJoinData(transportInsert));
        sheetGekauft.getRange(indexFuerNeueDaten, 17, anzahlZeilenToInsert, 1).setValues(convertIn2dArrayAndJoinData(transportBesonderheitInsert));
    }

    // Stand befüllen
    printStandInZelle("B2", sheetGekauft);
    Browser.msgBox(anzahlZeilenToInsert + " Datensätze eingefügt/ergänzt");
}
