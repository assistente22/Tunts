const https = require('https');
const excel = require('exceljs');

const filename = 'tunts.xlsx'

const wb = new excel.Workbook();
const ws = wb.addWorksheet('sheet');

//configuring cells

ws.mergeCells('A1:D1') //merging title cell

const columnA = ws.getColumn('A') //name column
const columnB = ws.getColumn('B') //capital column
const columnC = ws.getColumn('C') //area column
const columnD = ws.getColumn('D') //currencies column
const cellTitle = ws.getCell('A1') //title cell

columnA.alignment = {wrapText: true}
columnB.alignment = {wrapText: true}
columnC.alignment = {wrapText: true}
columnD.alignment = {wrapText: true}

columnC.numFmt = '0,0.00'
  
cellTitle.font = { 
    color: '#4F4F4F',
    size: 16,
    bold: true
}

cellTitle.alignment = {
    vertical: 'middle',
    horizontal:'center',
}

cellTitle.value = 'Countries List'


ws.getRow(2).font = {
    color: {argb: "#808080"},
    size: 12,
    bold: true
}

ws.getCell('A2').value = 'Name'
ws.getCell('B2').value = 'Capital'
ws.getCell('C2').value = 'Area'
ws.getCell('D2').value = 'Currencies'
columnA.width = 15
columnB.width = 15
columnC.width = 15
columnD.width = 15


let url = "https://restcountries.com/v3.1/all";

https.get(url,(res) => { //http request to api
    let body = "";

    res.on("data", (chunk) => {
        body += chunk;
    });

    res.on("end", () => {
        try {
            let json = JSON.parse(body);

            for(let i = 3; i< Object.keys(json).length +3; i++){ //parsing the json and setting values on the cells
                ws.getCell('A' + i).value = json[i-3]['name']['common']

                try{ //if the value is undefined insert - in the cell
                    ws.getCell('B' + i).value = Object.values(json[i-3]['capital'])[0]
                }
                catch(err){
                    ws.getCell('B' + i).value = '-'
                }

                ws.getCell('C' + i).value = json[i-3]['area']

                try{
                    let currenciesString = ''
                    for(let o = 0; o< Object.keys(json[i-3]['currencies']).length; o++){ // formatting the currencies
                        if(o == Object.keys(json[i-3]['currencies']).length-1){currenciesString = currenciesString + Object.keys(json[i-3]['currencies'])[o]}
                        else{ currenciesString = currenciesString + Object.keys(json[i-3]['currencies'])[o] + ", "}
                    }
                   ws.getCell("D" + i).value =  currenciesString
                }
                catch(err){
                    ws.getCell("D" + i).value = '-'
                }
            }

            wb.xlsx.writeFile(filename) //saving the xlsx file


            .then(() => {
                console.log('done!')
            })
            .catch(err => {
               console.log(err.message)
            })
            
        } catch (error) {
            console.error(error.message);
        };
    });

}).on("error", (error) => {
    console.error(error.message);
});
