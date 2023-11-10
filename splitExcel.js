function splitExcel() {
    const input = document.getElementById('excelFile');
    const file = input.files[0];

    if (!file) {
        alert('Please select an Excel file.');
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0]; // Assuming you want to split the first sheet.
        const worksheet = workbook.Sheets[sheetName];
        
        // Extract the data as an array of objects
        const dataArr = XLSX.utils.sheet_to_json(worksheet);

        // Group data by name from a filtered column (e.g., column A)
        const groupedData = dataArr.reduce((result, row) => {
            const name = row['Name']; // Adjust 'Name' to the actual header in your Excel file
            if (!result[name]) {
                result[name] = [];
            }
            result[name].push(row);
            return result;
        }, {});

        // Save each group as a separate Excel file
        for (const name in groupedData) {
            const newWorkbook = XLSX.utils.book_new();
            const newWorksheet = XLSX.utils.json_to_sheet(groupedData[name]);
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
            const outputData = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'blob' });

            // Create a Blob and save the file
            const blob = new Blob([outputData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${name}_output.xlsx`;
            a.click();
            window.URL.revokeObjectURL(url);
        }

        alert('Excel file has been split into separate files based on the filtered column.');
    };

    reader.readAsArrayBuffer(file);
}

