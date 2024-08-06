const systems = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
const gridContainer = document.getElementById('grid-container');
document.addEventListener('DOMContentLoaded', () => {
   fetch('data.xlsx')
       .then(response => response.arrayBuffer())
       .then(data => {
           const workbook = XLSX.read(data, { type: 'array' });
           const sheet = workbook.Sheets[workbook.SheetNames[0]];
           const incidents = XLSX.utils.sheet_to_json(sheet, { raw: true });
           displayGrid(incidents);
       });
});
function displayGrid(incidents) {
   const currentDate = new Date();
   systems.forEach(system => {
       const incident = incidents.find(incident => incident['System Name'] === system);
       let daysSince = 'No incidents';
       let lastIncidentDate = 'N/A';
       let smileyFace = 'green-smiley.png';
       if (incident) {
           const incidentDate = new Date(incident['Date']);
           const diffTime = currentDate - incidentDate;
           const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
           daysSince = `${diffDays} days`;
           lastIncidentDate = incidentDate.toISOString().split('T')[0];
           if (diffDays <= 14) {
               smileyFace = 'red-smiley.png';
           } else if (diffDays <= 30) {
               smileyFace = 'amber-smiley.png';
           }
       }
       const gridItem = document.createElement('div');
       gridItem.className = 'grid-item';
       gridItem.innerHTML = `
<h2>${system}</h2>
<img src="${smileyFace}" alt="Smiley Face" class="smiley-face">
<div class="counter">Days since last incident: ${daysSince}</div>
<div class="last-incident">Last incident: ${lastIncidentDate}</div>
       `;
       gridContainer.appendChild(gridItem);
   });
}
