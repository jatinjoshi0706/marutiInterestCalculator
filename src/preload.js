const { ipcRenderer } = require("electron");
const XLSX = require("xlsx");

document.addEventListener("DOMContentLoaded", function () {
  const fileSelector1 = document.querySelector("#file-input1");
  const fileSelector2 = document.querySelector("#file-input2");

  fileSelector1.addEventListener("change", (e) => {
    const filePath1 = e.target.files[0].path;
    ipcRenderer.send("file-selected1", filePath1);
    console.log(filePath1);
  });

  fileSelector2.addEventListener("change", (e) => {
    const filePath2 = e.target.files[0].path;
    ipcRenderer.send("file-selected2", filePath2);
    console.log(filePath2);
  });


  ipcRenderer.on("dataForExcelObj", (event, data) => {
    populateTable(data);
  });

  ipcRenderer.on("data-error", (event, errorMessage) => {
    console.error(errorMessage);
  });

  function populateTable(data) {
    const table = document.querySelector("table");
    table.innerHTML = "";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    for (const key in data[0]) {

      const th = document.createElement("th");
      th.innerText = key;
      headerRow.appendChild(th);

    }

    thead.appendChild(headerRow);
    table.appendChild(thead);
    const tbody = document.createElement("tbody");
    data.forEach((row) => {
      const tr = document.createElement("tr");
      for (const key in row) {
        const td = document.createElement("td");
        if (key === "Last Challan Date" || key === "Last Payment Date") {
          if (row[key] == 0) {
            td.innerText = "-";
          } else {
            const parseDueDate = XLSX.SSF.parse_date_code(row[key]);
            const dueDateMonth = parseDueDate.m;
            const dueDateDay = parseDueDate.d;
            const dueDateYear = parseDueDate.y;
            const dueDate = new Date(dueDateYear, dueDateMonth - 1, dueDateDay);

            td.innerText = dueDate.toLocaleDateString();
          }
        } else {
          td.innerText = row[key];
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }
});