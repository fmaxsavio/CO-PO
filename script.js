const marks = [4, 4, 4, 4, 4, 26, 26, 28];
let processedWorkbook;

function processExcel() {
    const fileInput = document.getElementById("uploadFile").files[0];
    if (!fileInput) {
        alert("Please upload an Excel file.");
        return;
    }

    const reader = new FileReader();
    reader.readAsArrayBuffer(fileInput);
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Identify last row with data in column C
        let lastRow = findLastRow(sheet, "C");

        // Process each value in column C from C3 onwards
        for (let row = 3; row <= lastRow; row++) {
            let cellRef = "C" + row;
            let cell = sheet[cellRef];

            if (!cell || isNaN(cell.v)) continue; // Skip invalid or empty cells

            const inputMarks = parseInt(cell.v);
            const splitMarks = splitMarksFunction(inputMarks);
            const columns = ["D", "E", "F", "G", "H", "I", "J", "K"];

            splitMarks.forEach((val, index) => {
                const newCellRef = columns[index] + row;
                sheet[newCellRef] = { t: "n", v: val }; // Ensure numeric data type
            });
        }

        // Explicitly update the range to reflect new cells
        sheet["!ref"] = `A1:K${lastRow}`;

        // Save processed workbook
        processedWorkbook = workbook;
        document.getElementById("downloadBtn").style.display = "inline";
        document.getElementById("status").innerText = "Processing complete! Click Download.";
    };
}

function downloadExcel() {
    if (!processedWorkbook) return;
    const wbout = XLSX.write(processedWorkbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Output.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Function to split marks correctly
function splitMarksFunction(input) {
    let remaining = input;
    let splitValues = [0, 0, 0, 0, 0, 0, 0, 0]; // Corresponds to D-K

    // Assign marks to columns D-H (Max 4 each)
    for (let i = 0; i < 5; i++) {
        if (remaining >= 4) {
            splitValues[i] = 4;
            remaining -= 4;
        } else {
            splitValues[i] = remaining;
            remaining = 0;
        }
    }

    // Assign marks to columns I-J (Max 26 each)
    for (let i = 5; i < 7; i++) {
        if (remaining >= 26) {
            splitValues[i] = 26;
            remaining -= 26;
        } else {
            splitValues[i] = remaining;
            remaining = 0;
        }
    }

    // Assign remaining marks to column K (Max 28)
    splitValues[7] = remaining; // Whatever is left goes to K

    return splitValues;
}

// Function to determine the last row in column C
function findLastRow(sheet, column) {
    let lastRow = 2; // Start at C3 (index 2)
    while (true) {
        let cellRef = column + (lastRow + 1);
        if (!sheet[cellRef] || isNaN(sheet[cellRef].v)) break;
        lastRow++;
    }
    return lastRow;
}
