const XLSX = require('xlsx');
const fs = require('fs');
const wget = require('wget-improved');
const Client = require('ftp');

const SRC = 'http://b2b.kare.de/KARE_Articel_delivery-dates_selling-unit-availability.xls';
const INPUT = './input/KARE_Articel_delivery-dates_selling-unit-availability.xls';
const OUTPUT_FILENAME = 'KARE_Articel_delivery-dates_selling-unit-availability.csv';
const OUTPUT = './output/' + OUTPUT_FILENAME;
const FTP_HOST = 'csv.kare.com.co';
const FTP_PATH = '/FechasAlemania/';
const USERNAME = 'redfred@csv.kare.com.co';
const PASSWORD = 'r3dfr3d01';

function getSpreadsheet() {
    const download = wget.download(SRC, INPUT, {});
    download.on('error', function(err) {
        console.log('There was an error downloading the spreadsheet: ', err);
    });
    download.on('end', function(output) {
        console.log('Spreadsheet has been downloaded.');
        translateSpreadsheetToCSV();
    });
}

function translateSpreadsheetToCSV() {
    const workbook = XLSX.readFile(INPUT);
    
    var result = [];
    workbook.SheetNames.forEach(function(sheetName) {
        var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
        if(csv.length > 0){
            result.push(csv);
        }
    });
    
    saveCSV(result.join("\n"));
}

function saveCSV(output) {
    if (output) {
        fs.writeFile(OUTPUT, output, function(err) {
            if (err) {
                console.log("There was an error translating and saving the CSV file: ", err);
            } else {
                console.log("File was successfully translated and saved.");
                transferCSVtoServer();
            }
        });
    }
}

function transferCSVtoServer() {
    const client = new Client();
    client.on('ready', function() {
        client.put(OUTPUT, FTP_PATH + OUTPUT_FILENAME, function(err) {
            if (err) {
                console.log("Error transferring CSV file to server: ", err);
            } else {
                console.log("CSV file transferred successfully.");
                client.end();
            }
        });
    });
    
    client.connect({
        host: FTP_HOST,
        user: USERNAME,
        password: PASSWORD
    });
}

getSpreadsheet();