// Stand 24.03.23


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

            let resortsForGegenstand = gegenstandZuResorts[gegenstandName];
            if (!resortsForGegenstand) {
                resortsForGegenstand = [];
            }
            resortsForGegenstand.push(resortZuordnung);
            gegenstandZuResorts[gegenstandName] = resortsForGegenstand;
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
                console.warn("Keinen Eintrag in der ResortlisteGesamt gefunden für Gegenstand: " + gegenstandName);
            }
        }
    }

    // Stand eintragen
    printStandInZelle("C2", sheetGesamt);

    Browser.msgBox("Resortzuteilung erfolgreich");
}
