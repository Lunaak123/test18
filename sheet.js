let data = [];
let filteredData = [];

// Load the Excel sheet
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: "NULL" }); // Assign "NULL" for null values
        filteredData = [...data];
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Display the Excel sheet data
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');

    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });

    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Apply operations based on user input
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());

    filteredData = data.filter(row => {
        const isPrimaryNull = row[primaryColumn] === "NULL";

        const columnChecks = operationColumns.map(col => {
            return operation === 'null' ? row[col] === "NULL" : row[col] !== "NULL";
        });

        return operationType === 'and'
            ? !isPrimaryNull && columnChecks.every(check => check)
            : !isPrimaryNull && columnChecks.some(check => check);
    });

    displaySheet(filteredData);
}

// Download the file in the selected format
function downloadFile() {
    const filename = document.getElementById('filename').value || 'downloaded_file';
    const fileFormat = document.getElementById('file-format').value;

    if (fileFormat === 'xlsx') {
        const worksheet = XLSX.utils.json_to_sheet(filteredData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, `${filename}.xlsx`);
    } else if (fileFormat === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
    } else if (fileFormat === 'pdf') {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        const tableColumn = Object.keys(filteredData[0]);
        const tableRows = [];

        filteredData.forEach(row => {
            const rowData = [];
            tableColumn.forEach(column => {
                rowData.push(row[column]);
            });
            tableRows.push(rowData);
        });

        doc.autoTable(tableColumn, tableRows);
        doc.save(`${filename}.pdf`);
    } else if (fileFormat === 'jpeg' || fileFormat === 'jpg') {
        html2canvas(document.getElementById('sheet-content')).then(canvas => {
            const imgData = canvas.toDataURL('image/jpeg');
            const link = document.createElement('a');
            link.href = imgData;
            link.download = `${filename}.jpg`;
            link.click();
        });
    } else {
        alert('Unsupported format!');
    }
}

// Event Listeners
document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'block';
});
document.getElementById('confirm-download').addEventListener('click', downloadFile);
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Load the Excel file (replace 'path/to/your/excel-file.xlsx' with your actual file path)
loadExcelSheet('path/to/your/excel-file.xlsx');
