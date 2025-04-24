  function updateDateTime() {
      const now = new Date();
      document.getElementById("dayTime").innerText =
        now.toLocaleTimeString('en-US') + " | " +
        now.toLocaleDateString('en-US', { weekday: 'long' });
      document.getElementById("date").innerText = now.toLocaleDateString();
    }
    setInterval(updateDateTime, 1000);

    function createDropdown(value = "") {
      const options = ["Sales", "Refund", "Service", "Other"];
      const select = document.createElement("select");
      options.forEach(opt => {
        const option = document.createElement("option");
        option.value = option.text = opt;
        if (opt === value) option.selected = true;
        select.appendChild(option);
      });
      return select;
    }

  // Monitor input changes in the table
  document.addEventListener('input', function (e) {
    const td = e.target;
    if (td.tagName === 'TD' && td.isContentEditable) {
      const row = td.parentElement;
      recalculateTotal(row);
    }
  });

  function recalculateTotal(row) {
    const gross = parseFloat(row.cells[3].innerText) || 0;
    const netVat = parseFloat(row.cells[4].innerText) || 0;
    const withTax = parseFloat(row.cells[5].innerText) || 0;

    const total = gross - netVat - withTax;
    row.cells[6].innerText = total.toFixed(2);
  }


    function addRow(company = "", invoice = "", gross = "", netVat = "", withTax = "", total = "", user = "", date = "", time = "", year = "", type = "Sales") {
      const row = document.createElement("tr");

      // Add dropdown for type
      const dropdownCell = document.createElement("td");
      dropdownCell.appendChild(createDropdown(type));
      row.appendChild(dropdownCell);

      [company, invoice, gross, netVat, withTax, total, user, date, time, year].forEach(text => {
        const cell = document.createElement("td");
        cell.contentEditable = true;
        cell.textContent = text;
        row.appendChild(cell);
      });

      const actionCell = document.createElement("td");
      actionCell.innerHTML = `<button class="delete-btn" onclick="deleteRow(this)">Delete</button>`;
      row.appendChild(actionCell);

      document.querySelector("#dataTable tbody").appendChild(row);
    }

    function deleteRow(button) {
      const row = button.closest("tr");
      row.remove();
    }

    function exportToExcel() {
      const table = document.getElementById('dataTable');
      const cloned = table.cloneNode(true);

      // Convert dropdowns to text
      Array.from(cloned.querySelectorAll("select")).forEach(select => {
        const cell = document.createElement("td");
        cell.textContent = select.value;
        select.parentNode.replaceWith(cell);
      });

      // Remove action column
      Array.from(cloned.rows).forEach(row => row.deleteCell(-1));

      const wb = XLSX.utils.table_to_book(cloned, { sheet: "Sales Data" });
      XLSX.writeFile(wb, "SalesDashboard.xlsx");
    }

    function importFromExcel() {
      const file = document.getElementById("excelFile").files[0];
      if (!file) return alert("Please select an Excel file.");

      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        rows.slice(1).forEach(row => {
          if (row.length >= 11) addRow(row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[0]);
        });
      };
      reader.readAsArrayBuffer(file);
    }