// Stand 15.04.23
const MATERIAL_GELIEHEN_START_ROW = 8;

function fillMaterialGeliehen() {
    var sheetGeliehen = SpreadsheetApp.getActive().getSheetByName('Material geliehen');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // Zu cachende Daten aus Bearbeitung: Anzahl geliefert / Kommentar Geliehen
    // Da Gegenstaende in der Regel mehrfach vorkommen wird als Key immer GegenstandName + Ausleiher verwendet
    var cacheNameAusleiherZuGeliefert = {};
    var cacheNameAusleiherZuKommentar = {};

    var rangeVorhandeneDatenGesamt = sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 2, MAX_ROWS, 8).getValues();
    rangeVorhandeneDatenGesamt.forEach(function (row) {

        let gegenstandName = row[0];
        let ausleiher = row[6];
        if (gegenstandName && ausleiher) {
            let key = gegenstandName + '_' + ausleiher;

            let geliefert = row[4];
            if (geliefert) {
                cacheNameAusleiherZuGeliefert[key] = geliefert;
            }

            let kommentar = row[7];
            if (kommentar) {
                cacheNameAusleiherZuKommentar[key] = kommentar;
            }
        }
    });
    let cacheSize = Object.keys(cacheNameAusleiherZuGeliefert).length + Object.keys(cacheNameAusleiherZuKommentar).length;
    console.log(cacheSize + " Einträge gecached");

    // Einträge mit "Resort besorgts selbst" aus Resortliste-Gesamt lesen
    var nameResortZuName = {};
    var nameResortZuResort = {};
    var nameResortZuAnzahl = {};
    var nameResortZuEinheit = {};
    var nameResortZuTransport = {};
    var nameResortZuTransportBesonderheit = {};

    var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(RESORT_GESAMT_LISTE_START_ROW, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 12).getValues();
    rangeResortlisteGesamtFiltered = rangeResortlisteGesamt.filter(row => row[8] == 'Resort');
    rangeResortlisteGesamtFiltered.forEach(function (row) {
        let gegenstandName = row[0];
        let resort = row[2];
        if (gegenstandName && resort) {
            let key = gegenstandName + '_' + resort;

            nameResortZuName[key] = gegenstandName;
            nameResortZuResort[key] = resort;

            let anzahl = row[3];
            mergeMap(nameResortZuAnzahl, key, anzahl);
            let einheit = row[4];
            mergeMap(nameResortZuEinheit, key, einheit);
            let transport = row[10];
            mergeMap(nameResortZuTransport, key, transport);
            let transportBesonderheit = row[11];
            mergeMap(nameResortZuTransportBesonderheit, key, transportBesonderheit);
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
    let anzahl = convertIn2dArrayAndSumData(Object.values(nameResortZuAnzahl));
    let einheit = convertIn2dArrayAndJoinData(Object.values(nameResortZuEinheit));
    let transport = convertIn2dArrayAndJoinData(Object.values(nameResortZuTransport));
    let transportBesonderheit = convertIn2dArrayAndJoinData(Object.values(nameResortZuTransportBesonderheit));

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
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 5, anzahlZeilen, 1).setValues(einheit);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 6, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 6, anzahlZeilen, 1).setValues(geliefert);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 8, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 8, anzahlZeilen, 1).setValues(resort);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 9, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 9, anzahlZeilen, 1).setValues(kommentar);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 10, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 10, anzahlZeilen, 1).setValues(transport);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 11, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 11, anzahlZeilen, 1).setValues(transportBesonderheit);

    printStandInZelle("B3", sheetGeliehen);
    Browser.msgBox("Material geliehen mit " + anzahlZeilen + " Zeilen befüllt.");

}
