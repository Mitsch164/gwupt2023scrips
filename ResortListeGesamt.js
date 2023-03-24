// Stand 24.03.23

const RESORT_LISTE_START_ROW = 8;
const GESAMT_LISTE_START_ROW = 18;

const SHEET_NAME_TO_RESORT_NAME = {
    'Resort Material Import': 'Material'

};

function fillResortListeGesamt() {
    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');
    var currentIndexGesamt = GESAMT_LISTE_START_ROW;

    for (const[sheetName, resortName]of Object.entries(SHEET_NAME_TO_RESORT_NAME)) {
        console.log('Starte Verarbeitung von Sheet: ' + sheetName);

        var sheetResort = SpreadsheetApp.getActive().getSheetByName(sheetName);
        var currentIndexResort = RESORT_LISTE_START_ROW;

        // Alle Zeilen schreiben bei denen ein Name ausgefüllt ist
        var nameGegenStandinZeile = sheetResort.getRange(currentIndexResort, 1).getValue();
        while (nameGegenStandinZeile) {

            // copy Name
            sheetGesamt.getRange(currentIndexGesamt, 2).setValue(nameGegenStandinZeile);

            // Fill Resortname
            sheetGesamt.getRange(currentIndexGesamt, 4).setValue(resortName);

            // copy restliche Spalten
            let restData = sheetResort.getRange(currentIndexResort, 2, 1, 11).getValues();
            sheetGesamt.getRange(currentIndexGesamt, 5, 1, 11).setValues(restData);

            // Indizes hochzählen und nächste Nameszeile befüllen
            currentIndexResort++;
            currentIndexGesamt++;

            nameGegenStandinZeile = sheetResort.getRange(currentIndexResort, 1).getValue();
        }

        console.log('Finished Resort ' + resortName + ' mit ' + (currentIndexGesamt - GESAMT_LISTE_START_ROW) + ' Zeilen');
    }

    // Stand befüllen
    printStandInZelle("C2", sheetGesamt);

    Browser.msgBox('Befüllung ResortListeGesamt mit ' + (currentIndexGesamt - GESAMT_LISTE_START_ROW) + ' Zeilen erfolgreich');
}
