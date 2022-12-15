// excel library
const XLSX = require('xlsx');
const axios = require('axios');
const dateFormat = require('dateformat');
var path = require('path');
var fs=require('fs');
const appRoot = path.resolve(__dirname);

var excelPath = appRoot + '/data';

if (!fs.existsSync(excelPath)){
	console.log("no dir ",excelPath);
	return;
}

var files=fs.readdirSync(excelPath);
// console.log('length:', files);

for(var i=0; i<files.length; i++){
	var strArray = files[i].split('-');
	var filename=path.join(excelPath, files[i]);

	// read excel file
	const workbook = XLSX.readFile(filename, { 'cellDates': true, 'dateNF': 'dd/mm/yyyy'});
	const sheet_name_list = workbook.SheetNames;
	var jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

	// get the company name from excel filename
	var companyName = strArray[0].trim();

	// reconfigurate the json data
	var temp = {};
	jsonData.forEach((element, index) => {
		temp.CompanyName = companyName;
		temp.EventID = element['Package Ref'];
		temp.Name = element['Variant'];
		temp.EventStartDate = dateFormat(element['Commences'], 'yyyy-mm-dd');
		temp.EventEndDate = dateFormat(element['Finishes'], 'yyyy-mm-dd');
		// send the post request
		axios.post('http://webhooks.smartreportz.com', temp)
		.then(function (response) {
			// handle success
			console.log(response);
		})
		.catch(function (error) {
			// handle error
			console.log(error);
		})
		.finally(function () {
			// always executed
		});
		console.log('json data:', temp);
		temp = {}
	});
}



