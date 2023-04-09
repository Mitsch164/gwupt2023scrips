// Stand 09.04.23


const SHEET_NAME_TO_RESORT_NAME = {
    'Resort Material Import': 'Material',
    'Resort Bar Import': 'Bar',
    'Resort Küche Import': 'Küche',
    'Resort ÖffArbeit Import': 'ÖffArbeit',
    'Resort NotfallMgmt Import': 'NotfallMgmt',
    'Resort PL Import': 'Projektleitung',
    'Resort Inhalt Import': 'Inhalt',
    'Resort Orga Import': 'Orga',
    'AK Jupfis': 'AK Jupfis',
    'AK Wölis': 'AK Wölis',
    'AK Pfadis': 'AK Pfadis',
    'AK Rover': 'AK Rover'
};

function fillResortListeGesamt() {
    var sheetGesamt = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');
    var currentIndexGesamt = RESORT_GESAMT_LISTE_START_ROW;
    var resortStatistik = "";

    for (const [sheetName, resortName] of Object.entries(SHEET_NAME_TO_RESORT_NAME)) {
        console.log('Starte Verarbeitung von Sheet: ' + sheetName);

        var sheetResort = SpreadsheetApp.getActive().getSheetByName(sheetName);
        var maxRowResort = RESORT_LISTE_START_ROW;

        // Letzte ausgefüllte Zeile ermitteln
        var nameGegenStandinZeile = sheetResort.getRange(maxRowResort, 2).getValue();
        while (nameGegenStandinZeile) {
            maxRowResort++;
            nameGegenStandinZeile = sheetResort.getRange(maxRowResort, 2).getValue();
        }

        var anzahlZeilen = maxRowResort - RESORT_LISTE_START_ROW;

        if (anzahlZeilen == 0) {
            console.log('Überspringe Leere Liste von Resort ' + resortName);
            continue;
        }

        // Fill GegenstandName
        var resorGegenstaende = sheetResort.getRange(RESORT_LISTE_START_ROW, 2, anzahlZeilen, 1).getValues();
        sheetGesamt.getRange(currentIndexGesamt, 2, anzahlZeilen, 1).setValues(resorGegenstaende);

        // Fill Resort name
        sheetGesamt.getRange(currentIndexGesamt, 4, anzahlZeilen, 1).setValue(resortName);

        // copy restliche Spalten
        let restData = sheetResort.getRange(RESORT_LISTE_START_ROW, 3, anzahlZeilen, 13).getValues();
        sheetGesamt.getRange(currentIndexGesamt, 5, anzahlZeilen, 13).setValues(restData);

        // kopierte Zeilen in Resortliste markieren
        sheetResort.getRange(RESORT_LISTE_START_ROW, 1, MAX_ROWS, 1).clearContent();
        sheetResort.getRange(RESORT_LISTE_START_ROW, 1, anzahlZeilen, 1).setValue('x');

        // Indizes hochzählen
        currentIndexGesamt = currentIndexGesamt + anzahlZeilen;

        // Statistik je Resort schreiben
        resortStatistik = resortStatistik + "Resort " + resortName + ": " + anzahlZeilen + " Zeilen \\n";

        console.log('Verarbeitung Resort ' + resortName + ' mit ' + anzahlZeilen + ' Zeilen abgeschlossen');
    }

    // Stand befüllen
    printStandInZelle("C2", sheetGesamt);

    let msg = 'Befüllung ResortListeGesamt mit ' + (currentIndexGesamt - RESORT_GESAMT_LISTE_START_ROW) + ' Zeilen erfolgreich \\n' + resortStatistik;
    Browser.msgBox(msg);
}
