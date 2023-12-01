async function convertAndDownload() {
    try {
        const fileInput = document.getElementById('fileInput');
        const file = fileInput.files[0];

        if (!file) {
            alert('Please select a text file.');
            return;
        }

        const textData = await readFile(file);

        // Convert text to Excel
        const convertedWorkbook = await convertTextToExcel(textData);

        // Save the converted Excel file
        const blob = await convertedWorkbook.xlsx.writeBuffer();
        const blobObject = new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blobObject);
        downloadLink.download = 'converted_report.xlsx';
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    } catch (error) {
        console.error(error);
        alert('Error converting text to Excel.');
    }
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => resolve(event.target.result);
        reader.onerror = (error) => reject(error);
        reader.readAsText(file);
    });
}

// Function to convert text to Excel
async function convertTextToExcel(textData) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet 1');

    // Implement your logic to parse and arrange data here
    // For simplicity, let's assume the text data has rows and columns separated by spaces
    const rows = textData.split('\n');

    let dividerLine = 0;
    let currentRow = 1;

    rows.forEach((row, rowIndex) => {
        // Skip lines containing only dashes
        if (/^-+$/.test(row.trim())) {
            ++dividerLine;
            return;
        }

        if (dividerLine !== 1) {
            if (dividerLine % 4 !== 2) {
                return;
            }
        }

        let columns = row.split(/\s+/).filter(Boolean);

        if (dividerLine !== 1) {
            columns = processNames(columns);
            columns = processPhone(columns);
            columns = convertToNumber(columns);
        }

        columns.forEach((column, columnIndex) => {
            worksheet.getCell(currentRow + 1, columnIndex + 1).value = column;
        });

        if (currentRow === 1) {
            ++currentRow;
        }
        ++currentRow;
    });

    return workbook;
}


function processNames(columns) {
    let inputString = columns[1];
    let result = inputString.match(/^(\d+)(.*)$/).slice(1);
    let newPosition = 1; // the position where you want to insert the new element
    if (result[1] !== "") {
        columns.splice(newPosition, 1, result[0], result[1]);
    }
    return columns;
}

function processPhone(columns) {
    let index = searchNumberCell(columns);
    if (index === -1) {
        return;
    }

    let inputString = columns[index];

    // Use a regular expression to match the closing parenthesis followed by any characters
    let regex = /(.*?\))/;

    // Use the match function with the regular expression to extract matches
    let matches = inputString.match(regex);

    let result = matches ? [matches[0], inputString.slice(matches[0].length)] : [];

    if (result[1] !== "") {
        columns.splice(index, 1, result[0], result[1]);
    }

    let startIndex = 2;
    let startIndexSecond = 3;
    columns = mergeNames(columns, startIndex, index);
    columns = mergeNames(columns, startIndexSecond, getLastTextIndex(columns));

    return columns;
}

function searchNumberCell(columns) {
    for (let i = 0; i < columns.length; i++) {
        const element = columns[i];
        if (element.includes('(') && element.includes(')')) {
            return i; // Return the index if both opening and closing brackets are found
        }
    }
    // Return -1 if no such element is found
    return -1;
}

function mergeNames(arr, startIndex, endIndex) {
    if (startIndex < 0 || endIndex >= arr.length || startIndex > endIndex) {
        console.log("Invalid indices provided.");
        return arr;
    }

    const mergedElement = arr.slice(startIndex, endIndex + 1).join(' ');
    arr.splice(startIndex, endIndex - startIndex + 1, mergedElement);

    return arr;
}

function getLastTextIndex(arr) {
    for (let i = arr.length - 1; i >= 0; i--) {
        if (arr[i].trim() !== '' && isNaN(arr[i])) {
            // Check if the element is not empty and not a number
            return i;
        }
    }
    return -1; // Return -1 if no element with text is found
}

function convertToNumber(inputArray) {
    const numericArray = inputArray.map(value => {
        const numericValue = parseFloat(value);
        return isNaN(numericValue) ? value : numericValue;
    });

    return numericArray;
}