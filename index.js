const { program } = require('commander');
if(typeof require !== 'undefined') XLSX = require('xlsx');
var jmespath = require('jmespath');

program
	.option('-d, --debug', 'output extra debugging')
	.option('-f, --file-xlsx <type>', 'input excel file path');

program.parse(process.argv);

if (program.fileXlsx) {
	console.log(`xlsx file path: ${program.fileXlsx}`);

	var workbook = XLSX.readFile(program.fileXlsx);
	var ia_sheet_name = 'IA Layout';
	var t4_sheet_name = 'T4 redirect pages';
	var ia_worksheet = workbook.Sheets[ia_sheet_name];
	var t4_worksheet = workbook.Sheets[t4_sheet_name];

	let ia_json_xlsx = XLSX.utils.sheet_to_json(ia_worksheet);
	let t4_json_xlsx = XLSX.utils.sheet_to_json(t4_worksheet);
	let search_result = jmespath.search(ia_json_xlsx, '[?"Row Type"==`IA Row`]');
	let csv_result = []; //the main results array

	console.log("Transforming 'IA Layout' sheet")
	for (const row of search_result){
		//console.log(row);
		// Parse Lx columns. Can't use the previous key from `IA Level` as it doesn't always exist
		let title_key = "";
		for (let i = 1; i < 10; i++) {
			title_key = "L" + i;
			if (row.hasOwnProperty(title_key)) {
				break;
			}
		}
		if (row[`INTERIM IA TRAILING SLASH FIXED`] == undefined || row[`END STATE IA TRAILING SLASH FIXED`] == undefined) {
			console.log(row);
			console.log("Interim IA:" + row[`INTERIM IA TRAILING SLASH FIXED`])
			console.log("Future IA:" + row[`END STATE IA TRAILING SLASH FIXED`])
		} else {
			let current_path = row[`INTERIM IA TRAILING SLASH FIXED`].replace(/\/$/, ""); //remove trailing slash
			let future_path = row[`END STATE IA TRAILING SLASH FIXED`].replace(/\/$/, ""); //remove trailing slash
			let row_json = { "current_ia": current_path, "future_ia": future_path, "title": row[title_key] };
			csv_result.push(row_json);
		}
	}

	console.log("Transforming 'T4 redirect pages' sheet")
	for (const row of t4_json_xlsx){
		//console.log(row);
		let current_path = row[`From`].replace(/\/$/, "");
		let future_path = row[`To`].replace(/\/$/, "");
		let row_json = { "current_ia": current_path, "future_ia": future_path };
		csv_result.push(row_json);
	}

	var csv_wb = XLSX.utils.book_new();
	var csv_ws = XLSX.utils.json_to_sheet(csv_result, {}); //opts to re-order columns: header:["current_ia","future_ia","title"]
	XLSX.utils.book_append_sheet(csv_wb, csv_ws, "FutureIA");
	XLSX.writeFile(csv_wb, 'ia-migrate-with-title.csv');

} else {
	console.log("Error: Missing xlsx spreadsheet path\n");
	console.log(program.helpInformation());
}
