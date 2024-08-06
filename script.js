const systems = ["Bob", "nia", "sim", "gar", "her", "Sid", "re2", "eacy", "75hk", "ash3", "ppp2", "qur"];
const gridContainer = document.getElementById('grid-container');
document.addEventListener('DOMContentLoaded', () => {
   fetch('data.xlsx')
       .then(response => response.arrayBuffer())
       .then(data => {
           const workbook = XLSX.read(data, {type: 'array'});
           const sheet = workbook.Sheets[workbook.SheetNames[0]];
           const incidents = XLSX.utils.sheet_to_json(sheet);
           displayGrid(incidents);
       });
});
function displayGrid(incidents) {
   const currentDate = new Date();
   systems.forEach(system => {
       const incident = incidents.find(incident => incident['System Name'] === system);
       let daysSince = 'No incidents';
       let lastIncidentDate = 'N/A';
       let smileyFace = 'green-smiley.jpg';
       if (incident) {
           // Correct date parsing using XLSX library utility functions
           const incidentDate = parseExcelDate(incident['Date']);
           const diffTime = currentDate - incidentDate;
           const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
           daysSince = `${diffDays} days`;
           lastIncidentDate = incidentDate.toISOString().split('T')[0];
           if (diffDays <= 14) {
               smileyFace = 'red-smiley.jpg';
           } else if (diffDays <= 30) {
               smileyFace = 'amber-smiley.jpg';
           }
       }
       const gridItem = document.createElement('div');
       gridItem.className = 'grid-item';
       gridItem.innerHTML = `
<h2>${system}</h2>
<img src="${smileyFace}" alt="Smiley Face" class="smiley-face">
<div class="counter">${daysSince}</div>
<div class="last-incident">Last incident: ${lastIncidentDate}</div>
       `;
       gridContainer.appendChild(gridItem);
   });
}
function parseExcelDate(excelDate) {
   // Excel serial date format
   const excelEpoch = new Date(1899, 11, 30); // Excel's epoch date
   const date = new Date(excelEpoch);
   date.setDate(date.getDate() + excelDate - 1);
   return date;
}
