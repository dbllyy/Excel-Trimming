let jsonData = [];  // Simpan data Excel dalam variabel global

document.getElementById('uploadExcel').addEventListener('change', function (event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Ambil sheet pertama
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Convert ke format JSON
        jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Render tabel berdasarkan data JSON
        renderTable();
    };

    reader.readAsArrayBuffer(file);
});

function renderTable() {
    const tableBody = document.querySelector('#excelTable tbody');
    tableBody.innerHTML = "";  // Bersihkan tabel sebelum menampilkan data baru

    jsonData.forEach((row, index) => {
        if (index === 0 || 
            row.includes("Status Terminal ATM") || 
            row.includes("Tanggal") || 
            row.includes("Jam Cetak") || 
            row.includes("User Cetak") || 
            row.includes("Profile") || 
            row.includes(":") || 
            row.includes("evn95") || 
            row.includes("Lokasi") || 
            row.every(cell => cell === "")) { // Mengabaikan baris kosong
            return;  // Abaikan baris ini dan lanjutkan ke baris berikutnya
        }

        const validRow = row.some(cell => {
            return typeof cell === 'string' && cell.trim() !== "" && !cell.match(/^[A-Za-z\s]*$/);
        });

        if (!validRow) {
            return; // Abaikan baris yang tidak memiliki data penting
        }

        const newRow = document.createElement('tr');

        row.forEach(cell => {
            const newCell = document.createElement('td');
            newCell.textContent = cell;
            newRow.appendChild(newCell);
        });

        tableBody.appendChild(newRow);
    });

    // Apply filters
    filterTable();
}

function filterTable() {
    const input = document.getElementById("searchInput").value.toLowerCase();
    const checkboxChecked = document.getElementById("depositCheckbox").checked;
    const table = document.getElementById("excelTable");
    const trs = table.getElementsByTagName("tr");

    for (let i = 1; i < trs.length; i++) {
        let tds = trs[i].getElementsByTagName("td");
        let resiCell = tds[4];
        let depositCell = tds[6];

        let showRow = true;

        if (checkboxChecked && depositCell) {
            let depositValue = depositCell.textContent.replace(/,/g, '');
            depositValue = parseFloat(depositValue);
            if (isNaN(depositValue) || depositValue > 50000000) {
                showRow = false;
            }
        }

        if (resiCell && !resiCell.textContent.toLowerCase().includes(input)) {
            showRow = false;
        }

        trs[i].style.display = showRow ? "" : "none";
    }
}

function exportToExcel() {
    const table = document.getElementById("excelTable");
    const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
    XLSX.writeFile(wb, "ExportedTable.xlsx");
}

function exportToPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    const headers = [["Profile", "Lokasi", "Status Koneksi", "Status Terminal", "Resi", "CashBox", "Deposit (RP)"]];
    const data = [];
    const trs = document.querySelectorAll("#excelTable tbody tr");

    trs.forEach(tr => {
        if (tr.style.display !== "none") {
            const rowData = Array.from(tr.getElementsByTagName("td")).map(td => td.textContent);
            data.push(rowData);
        }
    });

    doc.text("Status Terminal ATM", 14, 10);

    doc.autoTable({
        head: headers,
        body: data,
        startY: 20,
        styles: { fontSize: 8 },
    });

    doc.save("ExportedTable.pdf");
}
