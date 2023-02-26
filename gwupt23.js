// Stand 25.02.23
const MAX_ROWS = 500;


function resortzuteilungErmitteln() {

    var sheetGeamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // Resorts für jeden Gegenstand auf ResortlisteGesamt ermitteln
    var gegenstandZuResorts = {};
    var resortListeGesamtGegenstaende = sheetResortlisteKomplett.getRange(5, 1, MAX_ROWS).getValues();
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
    var gesamtListeGegenstaende = sheetGeamt.getRange(7, 1, MAX_ROWS).getValues();
    var gesamtListeResorts = sheetGeamt.getRange(7, 4, MAX_ROWS);

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
    var standFormatted = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.YYYY");
    var standCell = sheetGeamt.getRange("B2").getCell(1,1);
    standCell.setValue(standFormatted);

    Browser.msgBox("Resortzuteilung erfolgreich");
}
