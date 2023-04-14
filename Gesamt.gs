// Stand 14.04.23

function fillGesamt() {

    let anzahlZeilenBefuellt = groupGegenstaendeFromResortListeGesamt();

    resortzuteilungErmitteln();

    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    printStandInZelle("E1", sheetGesamt);

    Browser.msgBox("Gesamtliste mit " + anzahlZeilenBefuellt + " Zeilen befüllt.");
}

function groupGegenstaendeFromResortListeGesamt() {
    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // per Hand ausgefüllte Daten aus Gesamt  merken - Kategorie / gekauft / geliehen
    var cacheGegenstandZuKategorie = {};
    var cacheGegenstandZuGeliehen = {};
    var cachgeGegenstandZuGekauft = {};
    var cacheGegenstandZuAbzugParallelNutzung = {};

    var rangeVorhandeneDatenGesamt = sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 2, MAX_ROWS, 10).getValues();
    rangeVorhandeneDatenGesamt.forEach(function (row) {
        let gegenstandName = row[0];
        if (gegenstandName) {
            let kategorie = row[1];
            if (kategorie) {
                cacheGegenstandZuKategorie[gegenstandName] = kategorie;
            }
            let geliehen = row[3];
            if (geliehen) {
                cacheGegenstandZuGeliehen[gegenstandName] = geliehen;
            }
            let gekauft = row[4];
            if (gekauft) {
                cachgeGegenstandZuGekauft[gegenstandName] = gekauft;
            }
            let abzugParalleleNutzung = row[9];
            if (abzugParalleleNutzung) {
                cacheGegenstandZuAbzugParallelNutzung[gegenstandName] = abzugParalleleNutzung;
            }
        }
    });

    var gegenstandZuResortBesorgtsSelbst = {};
    var gegenstandZuKommentarResort = {};
    var gegenstandZuKommentarResortMaterial = {};
    var gegenstandZuEinheit = {};

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

            let einheit = row[4];
            mergeMap(gegenstandZuEinheit, gegenstandName, einheit);
            let kommentarResort = row[14];
            mergeMap(gegenstandZuKommentarResort, gegenstandName, kommentarResort);
            let kommentarMaterial = row[15];
            mergeMap(gegenstandZuKommentarResortMaterial, gegenstandName, kommentarMaterial);
        }
    });

    var gegenstandNamenResortListeGesamt = Object.keys(gegenstandZuKommentarResort);
    let anzahlZeilen = gegenstandNamenResortListeGesamt.length;
    console.log("Befülle Gesamtliste mit ", anzahlZeilen, " Zeilen");

    let gegenstandNamen = convertIn2dArray(Object.keys(gegenstandZuKommentarResort));
    let resortBesorgtsSelbst = convertIn2dArray(Object.values(gegenstandZuResortBesorgtsSelbst));
    let einheiten = convertIn2dArrayAndJoinData(Object.values(gegenstandZuEinheit));
    let kommentareResort = convertIn2dArrayAndJoinData(Object.values(gegenstandZuKommentarResort));
    let kommentareMaterial = convertIn2dArrayAndJoinData(Object.values(gegenstandZuKommentarResortMaterial));

    // Gecachte Daten aus Gesamtliste zuordnen
    let gegenstandNameZuKategorie = cacheWerteZuordnen(gegenstandNamenResortListeGesamt, cacheGegenstandZuKategorie);
    let kategorien = convertIn2dArray(Object.values(gegenstandNameZuKategorie));
    let gegenstandNameZuGeliehen = cacheWerteZuordnen(gegenstandNamenResortListeGesamt, cacheGegenstandZuGeliehen);
    let geliehen = convertIn2dArray(Object.values(gegenstandNameZuGeliehen));
    let gegenstandNameZuGekauft = cacheWerteZuordnen(gegenstandNamenResortListeGesamt, cachgeGegenstandZuGekauft);
    let gekauft = convertIn2dArray(Object.values(gegenstandNameZuGekauft));
    let gegenstandZuAbzugParalleleNutzung = cacheWerteZuordnen(gegenstandNamenResortListeGesamt, cacheGegenstandZuAbzugParallelNutzung);
    let abzugParalleleNutzung = convertIn2dArray(Object.values(gegenstandZuAbzugParalleleNutzung));

    // Daten in Gesamtliste schreiben
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 2, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 2, anzahlZeilen, 1).setValues(gegenstandNamen);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 3, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 3, anzahlZeilen, 1).setValues(kategorien);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 5, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 5, anzahlZeilen, 1).setValues(geliehen);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 6, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 6, anzahlZeilen, 1).setValues(gekauft);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 7, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 7, anzahlZeilen, 1).setValues(resortBesorgtsSelbst);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 11, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 11, anzahlZeilen, 1).setValues(abzugParalleleNutzung);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 15, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 15, anzahlZeilen, 1).setValues(einheiten);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 16, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 16, anzahlZeilen, 1).setValues(kommentareResort);

    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 17, MAX_ROWS, 1).clearContent();
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 17, anzahlZeilen, 1).setValues(kommentareMaterial);
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

    // Zuteilung in Gesamtliste eintragen
    sheetGesamt.getRange(GESAMT_LISTE_START_ROW, 4, MAX_ROWS, 1).clearContent();
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
