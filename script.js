document.addEventListener("DOMContentLoaded", () => {
    const tableContainer = document.getElementById("tableContainer");
    const loader = document.getElementById("loader"); // Get the loader element

    // URL of the Google Sheet published as XLSX
    const excelURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0f0gLQZ2jTCv8BBBnRXAEAXo1C3vEYDL9qDTh0hdrjgyzScUsidr0Um-NuBXJXda8FM_FRcCbfZaa/pub?output=xlsx";

    // Extract parameters from the URL
    const urlParams = new URLSearchParams(window.location.search);
    const line = urlParams.get("line") || "Summary_Line1"; // Default to Summary_Line1
    const number = parseInt(urlParams.get("number"), 10); // Parse the number parameter

    // Show the loader before fetching data
    loader.style.display = "block";

    // Fetch the Excel file and display its data
    fetch(excelURL)
        .then(response => response.arrayBuffer()) // Fetch as binary data
        .then(data => {
            const workbook = XLSX.read(data, { type: "array" }); // Parse Excel data

            // Check if the sheet exists
            if (!workbook.SheetNames.includes(line)) {
                tableContainer.innerHTML = `<p>Sheet "${line}" not found in the Excel file.</p>`;
                loader.style.display = "none"; // Hide the loader
                return;
            }

            const sheet = workbook.Sheets[line]; // Get the specified sheet
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convert to JSON (array of arrays)

            // If a number is specified, locate the row; otherwise, display the full table
            if (number && !isNaN(number)) {
                displayRowByNumber(jsonData, number);
            } else {
                displayTable(jsonData);
            }

            loader.style.display = "none"; // Hide the loader after data is loaded
        })
        .catch(error => {
            console.error("Error fetching Excel file:", error);
            tableContainer.innerHTML = "<p>Failed to load data. Please try again later.</p>";
            loader.style.display = "none"; // Hide the loader in case of an error
        });

    // Function to display a specific row by number
    function displayRowByNumber(data, number) {
        if (!data || data.length === 0) {
            tableContainer.innerHTML = "<p>No data found in the Excel file.</p>";
            loader.style.display = "none"; // Ensure loader is hidden
            return;
        }

        const headers = data[0]; // First row as headers
        const rows = data.slice(1); // All rows excluding the headers

        // Locate the row with the specified number in the first column
        const matchingRow = rows.find(row => parseInt(row[0], 10) === number);

        if (matchingRow) {
            const heading = document.createElement("h2");
            heading.textContent = `LINE ${line.includes("1") ? "1" : "2"}`; // Display only LINE x (1 or 2)
            tableContainer.appendChild(heading);

            // Display each field as a label-value pair
            headers.forEach((header, index) => {
                const fieldContainer = document.createElement("div");
                fieldContainer.classList.add("field-container");

                const label = document.createElement("span");
                label.classList.add("field-label");
                label.textContent = `${header}:`;

                const value = document.createElement("span");
                value.classList.add("field-value");

                // Modify the display value for the "status" column
                value.textContent = (header.toLowerCase() === "status" && matchingRow[index] === "DEAD CELL")
                    ? "OUT"
                    : matchingRow[index] || "N/A";

                fieldContainer.appendChild(label);
                fieldContainer.appendChild(value);
                tableContainer.appendChild(fieldContainer);
            });
        } else {
            tableContainer.innerHTML = `<p>No data found for Number: ${number} in sheet "${line}".</p>`;
        }

        loader.style.display = "none"; // Ensure loader is hidden when displaying the data
    }

    // Function to display the entire sheet in an HTML format (not used when number is specified)
    function displayTable(data) {
        if (!data || data.length === 0) {
            tableContainer.innerHTML = "<p>No data found in the Excel file.</p>";
            loader.style.display = "none"; // Hide the loader
            return;
        }

        const heading = document.createElement("h2");
        heading.textContent = `Data from ${line}`;
        tableContainer.appendChild(heading);

        // Display each field as a label-value pair for each row
        data.slice(1).forEach(row => {
            row.forEach((cell, index) => {
                const fieldContainer = document.createElement("div");
                fieldContainer.classList.add("field-container");

                const label = document.createElement("span");
                label.classList.add("field-label");
                label.textContent = `${data[0][index]}:`;

                const value = document.createElement("span");
                value.classList.add("field-value");

                // Modify the display value for the "status" column
                value.textContent = (data[0][index].toLowerCase() === "status" && cell === "DEAD CELL")
                    ? "OUT"
                    : cell || "N/A";

                fieldContainer.appendChild(label);
                fieldContainer.appendChild(value);
                tableContainer.appendChild(fieldContainer);
            });
        });

        loader.style.display = "none"; // Ensure loader is hidden after data is loaded
    }
});
