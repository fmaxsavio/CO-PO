const marks = [4,4,4,4,4,26,26,28];
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

        // Read value from C3
        const cellC3 = sheet["C3"];
        if (!cellC3 || isNaN(cellC3.v)) {
            alert("C3 is empty or invalid!");
            return;
        }
        const inputMarks = parseInt(cellC3.v);

        // Get split-up marks
        const splitMarks = MarkSplitUp(inputMarks);

        // Write values to D4 to K4
        const columns = ["D", "E", "F", "G", "H", "I", "J", "K"];
        splitMarks.forEach((val, index) => {
            const cellRef = columns[index] + "4";
            sheet[cellRef] = { v: val };
        });

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
    return out[0].split(",").map(Number);
}

function findSplitUp(input) {
    let out = "";
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
    return [splitValues.join(","), toCut, sum];
}

function randomIntFromInterval(min, max) {
    return Math.floor(Math.random() * (max - min + 1) + min);
}
