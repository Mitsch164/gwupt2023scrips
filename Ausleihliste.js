/**
 * Extrahiert die Werte für die Anzahlen aus dem Feld für private Ausleiher.
 */
function EXTRACTANZAHL(input, pattern, groupId) {
    var match = [];
    var extractedNumbers = [];
    if (input) {
        var regex = new RegExp(pattern, 'g');
        while (match = regex.exec(input)) {
            extractedNumbers.push(match[groupId]);
        }
    }
    var sum = 0;
    extractedNumbers.forEach(number => sum = sum + parseInt(number));
    return sum;
}
