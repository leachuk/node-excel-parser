if(typeof require !== 'undefined') XLSX = require('xlsx');
var jmespath = require('jmespath');

var workbook = XLSX.readFile('Master Full IA.xlsx');
var ia_sheet_name = 'IA Layout';
var t4_sheet_name = 'T4 redirect pages';
var ia_worksheet = workbook.Sheets[ia_sheet_name];
var t4_worksheet = workbook.Sheets[t4_sheet_name];

let ia_json_xlsx = XLSX.utils.sheet_to_json(ia_worksheet);
let t4_json_xlsx = XLSX.utils.sheet_to_json(t4_worksheet);
let search_result = jmespath.search(ia_json_xlsx, '[?"Row Type"==`IA Row`]');
let csv_result = [];

console.log("Transforming 'IA Layout' sheet")
for (const row of search_result){
	//console.log(row);
	let title_key = "L" + row[`IA Level`];
	let current_path = row[`INTERIM IA TRAILING SLASH FIXED`].replace(/\/$/, "");
	let future_path = row[`END STATE IA TRAILING SLASH FIXED`].replace(/\/$/, "");
	//let current_path_noslash = current_path.replace(/\/$/, "");
	let row_json = { "current_ia": current_path, "future_ia": future_path, "title": row[title_key] };
	csv_result.push(row_json);
}

console.log("Transforming 'T4 redirect pages' sheet")
for (const row of t4_json_xlsx){
	//console.log(row);
	let current_path = row[`From`].replace(/\/$/, "");
	let future_path = row[`To`].replace(/\/$/, "");
	let row_json = { "current_ia": current_path, "future_ia": future_path };
	csv_result.push(row_json);
}

//option: header:["current_ia","future_ia","title"]
var csv_wb = XLSX.utils.book_new();
var csv_ws = XLSX.utils.json_to_sheet(csv_result, {});
XLSX.utils.book_append_sheet(csv_wb, csv_ws, "FutureIA");
XLSX.writeFile(csv_wb, 'ia-migrate-with-title.csv');
