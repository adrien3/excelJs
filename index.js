var Excel = require('exceljs');

var workbook = new Excel.Workbook();

workbook.xlsx.readFile('ENGIE_TEMPLATE_manyPages.xlsx')
    .then(function() {
	 var worksheet = workbook.getWorksheet(3);

	 if(worksheet){
		// Add column
		var column = worksheet.getColumn("W");
		var id_new_column = 23;
		column._number = id_new_column;
		

		//Add info ligne.
		// iterate over all current cells in this column
		var row = worksheet.getRow(2);
		row.getCell(id_new_column).value = "Nom du config";

		if(worksheet._columns.length >= 22){
			worksheet._columns[22] = column;
		}else{
			worksheet._columns.push(column);
		}
	 }

	return workbook.xlsx.writeFile('new.xlsx');
});
