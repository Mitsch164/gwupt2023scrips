// Stand 08.04.23

function fillGesamt() {

    let anzahlZeilenBefuellt = groupGegenstaendeFromResortListeGesamt();

    resortzuteilungErmitteln();

    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    printStandInZelle("C2", sheetGesamt);

    Browser.msgBox("Gesamtliste mit " + anzahlZeilenBefuellt + " Zeilen befüllt.");
}

function groupGegenstaendeFromResortListeGesamt() {
    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // TODO Daten aus Gesamt vorladen merken - Kategorie / gekauft & geliehen
    // Anzahl Zeilen passt nicht zur Anzahl im Excel?

    var gegenstandZuResortBesorgtsSelbst = {};
    var gegenstandZuKommentarResort = {};
    var gegenstandZuKommentarResortMaterial = {};

    // Daten aus Resortliste gesamt lesen
    var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(5, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 16).getValues();
    rangeResortlisteGesamt.forEach(function (row) {
        let gegenstandName = row[0];

        if (gegenstandName && gegenstandName != 'AB HIER AUTOMATISCH') {
            let werBesorgts = row[8];
            if (werBesorgts == 'Resort') {
                gegenstandZuResortBesorgtsSelbst[gegenstandName] = 'x';
            } else {
                // Nur befüllen wenn noch kein x drinnensteht
                if (gegenstandZuResortBesorgtsSelbst[gegenstandName] != 'x') {
                    gegenstandZuResortBesorgtsSelbst[gegenstandName] = '';
                }
            }

            let kommentarResort = row[14];
            mergeMap(gegenstandZuKommentarResort, gegenstandName, kommentarResort);
            let kommentarMaterial = row[15];
            mergeMap(gegenstandZuKommentarResortMaterial, gegenstandName, kommentarMaterial);
        }
    });

    let anzahlZeilen = Object.keys(gegenstandZuKommentarResort).length;
    console.log("Befülle Gesamtliste mit ", anzahlZeilen, " Zeilen");

    console.log(JSON.stringify(gegenstandZuKommentarResort));

    let gegenstandNamen = convertIn2dArray(Object.keys(gegenstandZuKommentarResort));
    let resortBesorgtsSelbst = convertIn2dArray(Object.values(gegenstandZuResortBesorgtsSelbst));
    let kommentareResort = convertIn2dArrayAndJoinData(Object.values(gegenstandZuKommentarResort));
    let kommentareMaterial = convertIn2dArrayAndJoinData(Object.values(gegenstandZuKommentarResortMaterial));

    console.log(JSON.stringify(kommentareResort));

    // Daten in Gesamtliste schreiben
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 2, anzahlZeilen, 1).setValues(gegenstandNamen);
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 7, anzahlZeilen, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 7, anzahlZeilen, 1).setValues(resortBesorgtsSelbst);
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 14, anzahlZeilen, 1).setValues(kommentareResort);
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 15, anzahlZeilen, 1).setValues(kommentareMaterial);
    return anzahlZeilen;
}

function resortzuteilungErmitteln() {

    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // Resorts für jeden Gegenstand auf ResortlisteGesamt ermitteln
    var gegenstandZuResorts = {};
    var resortListeGesamtGegenstaende = sheetResortlisteKomplett.getRange(5, 2, MAX_ROWS).getValues();
    var resortListeGesamtResorts = sheetResortlisteKomplett.getRange(5, 4, MAX_ROWS).getValues();

    for (let row = 0; row < MAX_ROWS; row++) {
        let gegenstandName = resortListeGesamtGegenstaende[row][0];
        if (gegenstandName) {
            let resortZuordnung = resortListeGesamtResorts[row][0];
            mergeMap(gegenstandZuResorts, gegenstandName, resortZuordnung);
        }
    }
    //console.log(JSON.stringify(gegenstandZuResorts));

    // Zuteilung in Gesamtliste eintragen
    var gesamtListeGegenstaende = sheetGesamt.getRange(7, 2, MAX_ROWS).getValues();
    var gesamtListeResorts = sheetGesamt.getRange(7, 4, MAX_ROWS);

    for (let row = 0; row < MAX_ROWS; row++) {
        let gegenstandName = gesamtListeGegenstaende[row][0];
        if (gegenstandName) {
            let resortsFuerGegenstand = gegenstandZuResorts[gegenstandName];
            if (resortsFuerGegenstand) {
                let resortsJoined = resortsFuerGegenstand.join(",");

                // +1 weil range bei 1 anfängt zu zählen
                let resortCell = gesamtListeResorts.getCell(row + 1, 1);
                resortCell.setValue(resortsJoined);
            } else {
                throw new Error("Keinen Eintrag in der ResortlisteGesamt gefunden für Gegenstand: " + gegenstandName);
            }
        }
    }
}
