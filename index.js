const https = require('https');
const excel = require('exceljs');

const filename = 'tunts.xlsx'

const wb = new excel.Workbook();
const ws = wb.addWorksheet('sheet');

ws.mergeCells('A1:D1')

ws.getColumn('A').alignment = {wrapText: true}
ws.getColumn('B').alignment = {wrapText: true}
ws.getColumn('C').alignment = {wrapText: true}
ws.getColumn('D').alignment = {wrapText: true}

ws.getColumn('C').numFmt = '0,0.00'
  
ws.getCell('A1').font = { 
    color: '#4F4F4F',
    size: 16,
    bold: true
}


ws.getCell('A1').alignment = {
    vertical: 'middle',
    horizontal:'center',
}

ws.getCell('A1').value = 'Countries List'


ws.getRow(2).font = {
    color: '#808080',
    size: 12,
    bold: true
}

ws.getCell('A2').value = 'Name'
ws.getCell('B2').value = 'Capital'
ws.getCell('C2').value = 'Area'
ws.getCell('D2').value = 'Currencies'
ws.getColumn('A').width = 15
ws.getColumn('B').width = 15
ws.getColumn('C').width = 15
ws.getColumn('D').width = 15



let url = "https://restcountries.com/v3.1/all";

https.get(url,(res) => {
    let body = "";

    res.on("data", (chunk) => {
        body += chunk;
    });

    res.on("end", () => {
        try {
            let json = JSON.parse(body);

            for(let i = 3; i< Object.keys(json).length +3; i++){
                ws.getCell('A' + i).value = json[i-3]['name']['common']

                try{
                    ws.getCell('B' + i).value = Object.values(json[i-3]['capital'])[0]
                }
                catch(err){
                    ws.getCell('B' + i).value = '-'
                }

                ws.getCell('C' + i).value = json[i-3]['area']

                try{
                    let currenciesString = ''
                    for(let o = 0; o< Object.keys(json[i-3]['currencies']).length; o++){
                        if(o == Object.keys(json[i-3]['currencies']).length-1){currenciesString = currenciesString + Object.keys(json[i-3]['currencies'])[o]}
                        else{ currenciesString = currenciesString + Object.keys(json[i-3]['currencies'])[o] + ", "}
                    }
                   ws.getCell("D" + i).value =  currenciesString
                }
                catch(err){
                    ws.getCell("D" + i).value = '-'
                }
            }

            wb.xlsx.writeFile(filename)


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
