const XLSX = require('xlsx');
const fs = require('fs');
const express = require("express");
const app = express();
const port = 3000;
const upload = require('express-fileupload');
const co = require('co');
const generate = require('node-chartist');

app.use(upload());

app.post('/upload', async (req, res) => {

    var importedSheet = (XLSX.read(req.files.file.data, {type:'buffer'}));
    var sheet = importedSheet.Sheets['Company Totals'];
    var toJson = XLSX.utils.sheet_to_json(sheet);
    var AgencyName = [];
    var AgencyTotalQuote = [];
    var AgencyAppliedQuote = [];
    var branchName = ['Agency Totals'];
    var company = [];
    var quote = [];
    var apply = [];
    // Have to push the end HTML into an array, so I have loaded the CSS as the first element then the HTML will come in after
    var barArray = [fs.readFileSync(__dirname + '/public/main.html')];
    var count = 0;
    
    try{

        for(let i = 1; i < toJson.length; i++) {

            var agencyArr = (toJson[i]['Agency Totals:'].split(' '));

            if(agencyArr[0] !== 'Companies' && agencyArr[0] !== 'Branch:') {
                AgencyName.push(agencyArr[0]);
            }

            if(agencyArr[0] === 'Branch:') {
                branchName.push(agencyArr.join(''));
            }

            if(toJson[i]['__EMPTY'] !== 'Quote' && toJson[i]['__EMPTY'] !== undefined) {
                AgencyTotalQuote.push(toJson[i]['__EMPTY']);
            }

            if(toJson[i]['__EMPTY_1'] !== 'Apply' && toJson[i]['__EMPTY_1'] !== undefined) {
                AgencyAppliedQuote.push(toJson[i]['__EMPTY_1']);
            }
            if(agencyArr[0] === 'Branch:') {
                company.push(AgencyName);
                quote.push(AgencyTotalQuote);
                apply.push(AgencyAppliedQuote);
                AgencyTotalQuote = [];
                AgencyName = [];
                AgencyAppliedQuote = [];
            }
        }
            company.push(AgencyName);
            quote.push(AgencyTotalQuote);
            apply.push(AgencyAppliedQuote);

        for(let i = 0; i < company.length; i++) {
            var x = co(function * () {
                var data = {
                    labels: company[i],
                    series: [
                        {name: 'Quote', value: quote[i]},
                        {name: 'Apply', value: apply[i]}
                      ]
                  };

                  var options = {
                      height: 500,
                      width: 1800,
                    seriesBarDistance: 15
                  };
                var bar = `<h1>${branchName[i]}</h1>`
                bar += yield generate('bar', options, data); // => chart HTML
                bar += '<br><br><br><br>'
                barArray.push(bar);
            });

            (async function asyncCall() {
                var result = await x;
                if(result === undefined && count === company.length-1) {
                    res.send(barArray.join(''));
                }
                count++;
            })();

        }

    } catch(err) {
        console.log(err);
    }
    
});

app.listen(port, () => {
    console.log('server listening on port' + port);
})
