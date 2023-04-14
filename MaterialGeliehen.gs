// Stand 14.04.23
const MATERIAL_GELIEHEN_START_ROW = 8;

function fillMaterialGeliehen() {
    var sheetGeliehen = SpreadsheetApp.getActive().getSheetByName('Material geliehen');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // Zu cachende Daten aus Bearbeitung: Anzahl geliefert / Kommentar Geliehen
    // Da Gegenstaende in der Regel mehrfach vorkommen wird als Key immer GegenstandName + Ausleiher verwendet
    var cacheNameAusleiherZuGeliefert = {};
    var cacheNameAusleiherZuKommentar = {};

    var rangeVorhandeneDatenGesamt = sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 2, MAX_ROWS, 7).getValues();
    rangeVorhandeneDatenGesamt.forEach(function (row) {

        let gegenstandName = row[0];
        let ausleiher = row[5];
        if (gegenstandName && ausleiher) {
            let key = gegenstandName + '_' + ausleiher;

            let geliefert = row[3];
            if (geliefert) {
                cacheNameAusleiherZuGeliefert[key] = geliefert;
            }

            let kommentar = row[6];
            if (kommentar) {
                cacheNameAusleiherZuKommentar[key] = kommentar;
            }
        }
    });

    // Einträge mit "Resort besorgts selbst" aus Resortliste-Gesamt lesen
    var nameResortZuName = {};
    var nameResortZuResort = {};
    var nameResortZuAnzahl = {};
    var nameResortZuTransport = {};

    var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(RESORT_GESAMT_LISTE_START_ROW, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 11).getValues();
    rangeResortlisteGesamtFiltered = rangeResortlisteGesamt.filter(row => row[8] == 'Resort');
    rangeResortlisteGesamtFiltered.forEach(function (row) {
        let gegenstandName = row[0];
        let resort = row[2];
        if (gegenstandName && resort) {
            let key = gegenstandName + '_' + resort;

            nameResortZuName[key] = gegenstandName;
            nameResortZuResort[key] = resort;

            let anzahl = row[3];
            nameResortZuAnzahl[key] = anzahl;
            let transport = row[10]
                nameResortZuTransport[key] = transport;
        } else {
            throw new Error("Feld für Map Key nicht gesetzt in ResortListeGesamt");
        }
    });

    // TODO Daten aus Ausleihliste Zelte + Gesamt lesen und dazujoinen


    // Zu schreibende Daten aus "Resortliste gesamt" aufbereiten
    let keys = Object.keys(nameResortZuName);
    let anzahlZeilen = keys.length;
    console.log("Befülle Gesamtliste mit ", anzahlZeilen, " Zeilen");

    let gegenstandNamen = convertIn2dArray(Object.values(nameResortZuName));
    let resort = convertIn2dArray(Object.values(nameResortZuResort));
    let anzahl = convertIn2dArray(Object.values(nameResortZuAnzahl));
    let transport = convertIn2dArray(Object.values(nameResortZuTransport));

    // Gecachte Daten aus ursprünglicher Liste zuordnen
    let nameAusleiherZuGeliefert = cacheWerteZuordnen(keys, cacheNameAusleiherZuGeliefert);
    let geliefert = convertIn2dArray(Object.values(nameAusleiherZuGeliefert));
    let nameAusleiherZuKommentar = cacheWerteZuordnen(keys, cacheNameAusleiherZuKommentar);
    let kommentar = convertIn2dArray(Object.values(nameAusleiherZuKommentar));

    // Listeninhalte leeren und neue Daten schreiben
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 2, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 2, anzahlZeilen, 1).setValues(gegenstandNamen);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 4, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 4, anzahlZeilen, 1).setValues(anzahl);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 5, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 5, anzahlZeilen, 1).setValues(geliefert);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 7, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 7, anzahlZeilen, 1).setValues(resort);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 8, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 8, anzahlZeilen, 1).setValues(kommentar);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 9, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 9, anzahlZeilen, 1).setValues(transport);

    printStandInZelle("B3", sheetGeliehen);
    Browser.msgBox("Material geliehen mit " + anzahlZeilen + " Zeilen befüllt.");

}
