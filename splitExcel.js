function splitExcel() {
    const input = document.getElementById('excelFile');
    const file = input.files[0];

    if (!file) {
        alert('Please select an Excel file.');
        return;
    }

    const folderInput = document.getElementById('outputFolder');
    if (folderInput.files.length === 0) {
        alert('Please choose an output folder.');
        return;
    }

    const outputFolder = folderInput.files[0];

    const columnName = document.getElementById('columnName').value;

    if (!columnName) {
        alert('Please specify the column name for separation.');
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const dataArr = XLSX.utils.sheet_to_json(worksheet);

        const groupedData = dataArr.reduce((result, row) => {
            const name = row[columnName];
            if (!result[name]) {
                result[name] = [];
            }
            result[name].push(row);
            return result;
        }, {});

        for (const name in groupedData) {
            const newWorkbook = XLSX.utils.book_new();
            const newWorksheet = XLSX.utils.json_to_sheet(groupedData[name]);
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
            const outputData = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'blob' });

            const blob = new Blob([outputData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const outputFilePath = `${outputFolder.webkitRelativePath}/${name}_output.xlsx`;
            saveAs(blob, outputFilePath);
        }

        alert('Excel file has been split into separate files based on the specified column.');
    };

    reader.readAsArrayBuffer(file);
}

function saveAs(blob, fileName) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
}
