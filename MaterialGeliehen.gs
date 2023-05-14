// Stand 08.05.23
const MATERIAL_GELIEHEN_START_ROW = 8;
const AUSLEIHLISTE_START_ROW = 8;
const PRIVATE_AUSLEIHER_SPLIT_REGEX = new RegExp('\\(([a-zA-Z\\s]*:[\\d])\\)[,]?', 'g');

function fillMaterialGeliehen() {
    var sheetGeliehen = SpreadsheetApp.getActive().getSheetByName('Material geliehen');
    var sheetResortlisteKomplett = SpreadsheetApp.getActive().getSheetByName('Resortliste komplett');

    // 1) Zu cachende Daten aus Bearbeitung: Anzahl geliefert / Kommentar Geliehen
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

    // 2.1) Einträge mit "Resort besorgts selbst" aus Resortliste-Gesamt lesen
    var nameAusleiherZuGegenstandName = {};
    var nameAusleiherZuAusleiher = {};
    var nameAusleiherZuAnzahl = {};
    var nameAusleiherZuEinheit = {};
    var nameAusleiherZuTransport = {};
    var nameAusleiherZuTransportBesonderheit = {};

    var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(RESORT_GESAMT_LISTE_START_ROW, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 12).getValues();
    rangeResortlisteGesamtFiltered = rangeResortlisteGesamt.filter(row => row[8] == 'Resort');
    rangeResortlisteGesamtFiltered.forEach(function (row) {
        let gegenstandName = row[0];
        let resort = row[2];
        if (gegenstandName && resort) {
            let key = gegenstandName + '_' + resort;

            nameAusleiherZuGegenstandName[key] = gegenstandName;
            nameAusleiherZuAusleiher[key] = resort;

            let anzahl = row[3];
            mergeMap(nameAusleiherZuAnzahl, key, anzahl);
            let einheit = row[4];
            mergeMap(nameAusleiherZuEinheit, key, einheit);
            let transport = row[10];
            mergeMap(nameAusleiherZuTransport, key, transport);
            let transportBesonderheit = row[11];
            mergeMap(nameAusleiherZuTransportBesonderheit, key, transportBesonderheit);
        } else {
            throw new Error("Feld für Map Key nicht gesetzt in ResortListeGesamt");
        }
    });

    // 2.2) Daten aus Ausleihliste lesen und dazujoinen

    var sheetAusleihliste = SpreadsheetApp.getActive().getSheetByName('Ausleihliste Gesamt');

    // In Ausleihliste nicht enthaltene Daten für Transport + Besonderheit aus Resortliste Gesamt holen
    var gegenstandNameZuTransportLookup = {};
    var gegenstandNameZuTransportBesonderheitLookup = {};
    var gegenstandNamenAusleihliste = [];

    var gegenstandNamenAusleihlisteRange = sheetAusleihliste.getRange(AUSLEIHLISTE_START_ROW, 1, MAX_ROWS_RESORTLISTE_KOMPLETT, 3).getValues();
    gegenstandNamenAusleihlisteRange.forEach(function (row) {
        let gegenstandName = row[0];
        let geliehen = row[2];
        if (gegenstandName && geliehen == 'x') {
            gegenstandNamenAusleihliste.push(gegenstandName);
        }
    });

    var rangeResortlisteGesamt = sheetResortlisteKomplett.getRange(5, 2, MAX_ROWS_RESORTLISTE_KOMPLETT, 12).getValues();
    var resortListeGesamtFiltered = rangeResortlisteGesamt.filter(function (row) {
        return gegenstandNamenAusleihliste.includes(row[0]);
    });

    resortListeGesamtFiltered.forEach(function (row) {
        mergeMap(gegenstandNameZuTransportLookup, row[0], row[10]);
        mergeMap(gegenstandNameZuTransportBesonderheitLookup, row[0], row[11]);
    });

    // Daten aus Auswahlliste lesen und zu schreibende Daten ergänzen
    var headerInklusiveStammName = sheetAusleihliste.getRange(7, 1, 1, 25).getValues();

    var rangeAusleiherMitAnzahl = sheetAusleihliste.getRange(AUSLEIHLISTE_START_ROW, 1, MAX_ROWS_RESORTLISTE_KOMPLETT, 25).getValues();
    rangeAusleiherMitAnzahl.forEach(function (row) {
        let gegenstandName = row[0];
        let geliehen = row[2];
        if (gegenstandName && geliehen == 'x') {
            // Einzelne Stämme für Gegenstand durchgehen und zu schreibende Zeilen erzeugen
            for (let index = 8; index <= 22; index=index+2) {
                let anzahlAusgeliehen = row[index];
                if (anzahlAusgeliehen) {

                    let stamm = headerInklusiveStammName[0][index];
                    console.log(stamm);
                    let key = gegenstandName + "_" + stamm;

                    nameAusleiherZuGegenstandName[key] = gegenstandName;
                    nameAusleiherZuAusleiher[key] = stamm;
                    mergeMap(nameAusleiherZuAnzahl, key, anzahlAusgeliehen);

                    let einheit = row[4];
                    mergeMap(nameAusleiherZuEinheit, key, einheit);

                    let transport = gegenstandNameZuTransportLookup[gegenstandName];
                    mergeMap(nameAusleiherZuTransport, key, transport);

                    let transportBesonderheit = gegenstandNameZuTransportBesonderheitLookup[gegenstandName];
                    mergeMap(nameAusleiherZuTransportBesonderheit, key, transportBesonderheit);
                }
            }

            // private Ausleiher aufdröseln und zu schreibende Zeilen erzeugen
            let privateAusleiherString = row[24];
            if (privateAusleiherString) {
                let match = [];
                while (match = PRIVATE_AUSLEIHER_SPLIT_REGEX.exec(privateAusleiherString)) {
                    let ausleiherUndAnzahl = match[1];
                    let ausleiherUndAnzahlSplitted = ausleiherUndAnzahl.split(":");

                    let ausleiher = ausleiherUndAnzahlSplitted[0];
                    let anzahl = ausleiherUndAnzahlSplitted[1];
                    let key = gegenstandName + "_" + ausleiher;

                    nameAusleiherZuGegenstandName[key] = gegenstandName;
                    nameAusleiherZuAusleiher[key] = ausleiher;
                    mergeMap(nameAusleiherZuAnzahl, key, anzahl);

                    let einheit = row[4];
                    mergeMap(nameAusleiherZuEinheit, key, einheit);

                    let transport = gegenstandNameZuTransportLookup[gegenstandName];
                    mergeMap(nameAusleiherZuTransport, key, transport);

                    let transportBesonderheit = gegenstandNameZuTransportBesonderheitLookup[gegenstandName];
                    mergeMap(nameAusleiherZuTransportBesonderheit, key, transportBesonderheit);
                }
            }
        }
    });

    // 3) Zu schreibende Daten aus "Resortliste gesamt" aufbereiten und Daten schreiben
    let keys = Object.keys(nameAusleiherZuGegenstandName);
    let anzahlZeilen = keys.length;
    console.log("Befülle Gesamtliste mit ", anzahlZeilen, " Zeilen");

    let gegenstandNamen = convertIn2dArray(Object.values(nameAusleiherZuGegenstandName));
    let ausleiher = convertIn2dArray(Object.values(nameAusleiherZuAusleiher));
    let anzahl = convertIn2dArrayAndSumData(Object.values(nameAusleiherZuAnzahl));
    let einheit = convertIn2dArrayAndJoinData(Object.values(nameAusleiherZuEinheit));
    let transport = convertIn2dArrayAndJoinData(Object.values(nameAusleiherZuTransport));
    let transportBesonderheit = convertIn2dArrayAndJoinData(Object.values(nameAusleiherZuTransportBesonderheit));

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
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 8, anzahlZeilen, 1).setValues(ausleiher);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 9, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 9, anzahlZeilen, 1).setValues(kommentar);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 10, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 10, anzahlZeilen, 1).setValues(transport);

    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 11, MAX_ROWS, 1).clearContent();
    sheetGeliehen.getRange(MATERIAL_GELIEHEN_START_ROW, 11, anzahlZeilen, 1).setValues(transportBesonderheit);

    printStandInZelle("B3", sheetGeliehen);
    Browser.msgBox("Material geliehen mit " + anzahlZeilen + " Zeilen befüllt.");

}
