let jsonData = [];

document.getElementById('excelFile').addEventListener('change', function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    jsonData = XLSX.utils.sheet_to_json(worksheet);
    renderTable();
  };

  reader.readAsArrayBuffer(file);
});

function renderTable() {
  const table = document.getElementById('dataTable');
  table.innerHTML = '';
  if (!jsonData.length) return;

  const headers = Object.keys(jsonData[0]);
  let thead = '<tr>' + headers.map(h => `<th>${h}</th>`).join('') + '<th>Actions</th></tr>';
  table.innerHTML += thead;

  jsonData.forEach((row, i) => {
    let tr = '<tr>';
    headers.forEach(h => {
      tr += `<td contenteditable="true" oninput="editCell(${i}, '${h}', this.innerText)">${row[h]}</td>`;
    });
    tr += `<td><button onclick="deleteRow(${i})">Delete</button></td>`;
    tr += '</tr>';
    table.innerHTML += tr;
  });
}

function editCell(rowIndex, key, value) {
  jsonData[rowIndex][key] = value;
}

function deleteRow(index) {
  jsonData.splice(index, 1);
  renderTable();
}

function downloadJSON() {
  const blob = new Blob([JSON.stringify(jsonData, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'data.json';
  a.click();
  URL.revokeObjectURL(url);
}
