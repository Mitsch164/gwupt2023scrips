function resortzuteilungErmitteln() {

    var sheetGeamt = SpreadsheetApp.getActive().getSheetByName('Gesamt');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // Resorts für jeden Gegenstand auf ResortlisteGesamt ermitteln
    var gegenstandZuResorts = {};
    var resortListeGesamtGegenstaende = sheetResortlisteKomplett.getRange(5, 1, 500).getValues();
    var resortListeGesamtResorts = sheetResortlisteKomplett.getRange(5, 4, 500).getValues();

    for (let row = 0; row < 500; row++) {
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
    var gesamtListeGegenstaende = sheetGeamt.getRange(7, 1, 500).getValues();
    var gesamtListeResorts = sheetGeamt.getRange(7, 4, 500);

    for (let row = 0; row < 500; row++) {
        let gegenstandName = gesamtListeGegenstaende[row][0];
        if (gegenstandName) {
            let resortsFuerGegenstand = gegenstandZuResorts[gegenstandName];
            if (resortsFuerGegenstand) {
                let resortsJoined = resortsFuerGegenstand.join(",");
                let resortCell = gesamtListeResorts.getCell(row + 1, 1);
                resortCell.setValue(resortsJoined);
            } else {
                console.warn("Keinen Eintrag in der ResortlisteGesamt gefunden für Gegenstand: " + gegenstandName);
            }
        }
    }

    Browser.msgBox("Resortzuteilung erfolgreich");
}
