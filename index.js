var Excel = require('exceljs');

var workbook = new Excel.Workbook();

workbook.xlsx.readFile('ENGIE_TEMPLATE_manyPages.xlsx')
	.then(function () {
		var worksheet 	= workbook.getWorksheet(3);
		var index 		= workbook.getWorksheet(4);

		if (worksheet) {
			// Add column
			var column 			= worksheet.getColumn("W");
			var id_new_column 	= 22;
			var row 			= worksheet.getRow(2);
			var formule 		= ['index!$A$2:$A$3'];
			var begin_data 		= 3;
			var n_col_add 		= 40;
			var n_row_add 		= 25;


			for (let i = 1; i <= n_col_add; i++) {
				id_new_column = id_new_column + 1;
				column._number = id_new_column;
				row.getCell(id_new_column).value = "Nom du config " + i;

				for ( let j = begin_data; j < n_row_add; j++) {
					let newRow = worksheet.getRow(j);
					newRow.getCell(id_new_column).dataValidation = {
						type: 'list',
						allowBlank: true,
						formulae: formule
					};
				}

				if (worksheet._columns.length >= id_new_column) {
					worksheet._columns[id_new_column] = column;
				} else {
					worksheet._columns.push(column);
				}
			}

		}

		return workbook.xlsx.writeFile('new.xlsx');
	});
