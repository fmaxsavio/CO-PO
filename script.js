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
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Find the last row in column C
        let range = XLSX.utils.decode_range(sheet["!ref"]);
        let lastRow = range.e.r; // Default last row
        for (let row = 2; row <= lastRow; row++) { // Start from row index 2 (C3 in Excel)
            let cellRef = "C" + (row + 1);
            if (!sheet[cellRef] || isNaN(sheet[cellRef].v)) {
                lastRow = row;
                break;
            }
        }

        // Process each value in column C from C3 onwards
        for (let row = 2; row <= lastRow; row++) {
            let cellRef = "C" + (row + 1);
            let cell = sheet[cellRef];

            if (!cell || isNaN(cell.v)) continue; // Skip invalid or empty cells

            const inputMarks = parseInt(cell.v);
            const splitMarks = MarkSplitUp(inputMarks);
            const columns = ["D", "E", "F", "G", "H", "I", "J", "K"];

            splitMarks.forEach((val, index) => {
                const newCellRef = columns[index] + (row + 1);
                sheet[newCellRef] = { t: "n", v: val }; // Ensure numeric data type
            });
        }

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

function MarkSplitUp(input) {
    let out = findSplitUp(input);
    while (out[1] !== 0 || out[2] !== input) {
        out = findSplitUp(input);
    }
    return out[0];
}

function findSplitUp(input) {
    let toCut = 100 - input;
    let sum = 0;
    let splitValues = [];

    for (let i = 0; i < marks.length; i++) {
        if ((toCut - marks[i]) < 0 && toCut !== 0) {
            splitValues.push(marks[i] - toCut);
            sum += marks[i] - toCut;
            toCut = 0;
        } else {
            if (sum === input) {
                splitValues.push(0);
                toCut -= marks[i];
            } else if (toCut === 0 && sum < input) {
                splitValues.push(marks[i]);
                sum += marks[i];
            } else {
                let r = randomIntFromInterval(0, marks[i]);
                toCut -= r;
                sum += marks[i] - r;
                splitValues.push(marks[i] - r);
            }
        }
    }
    return [splitValues, toCut, sum];
}

function randomIntFromInterval(min, max) {
    return Math.floor(Math.random() * (max - min + 1) + min);
}
