function runCode() {
// Import the xlsx library
console.log('Importing xlsx library...');
const XLSX = require('xlsx');
console.log('Imported xlsx library.');

// Read the excel file
console.log('Reading excel file...');
const workbook = XLSX.readFile('leaderboard.xlsx');
console.log('Read excel file.');

// Get the first worksheet
console.log('Getting first worksheet...');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
console.log('Got first worksheet.');

// Convert the worksheet to an array of objects
console.log('Converting worksheet to array of objects...');
const data = XLSX.utils.sheet_to_json(worksheet);
console.log('Converted worksheet to array of objects.');

// Sort the data by time in ascending order
console.log('Sorting data by time in ascending order...');
data.sort((a, b) => a.time - b.time);
console.log('Sorted data by time in ascending order.');

// Display the leaderboard
console.log('Displaying leaderboard...');
data.forEach((row, index) => {
  console.log(`${index + 1}. ${row.name} - ${row.time}`);
});
console.log('Displayed leaderboard.');


let html = '<table>\n';
html += '<tr style="background-color: gold;"><th>Rank</th><th>Name</th><th>Time</th></tr>\n';
const row1 = data[0];
html += `<tr style="background-color: gold;"><td>1</td><td>${row1.name}</td><td>${row1.time}</td></tr>\n`;
const row2 = data[1];
const row3 = data[2];
html += `<tr style="background-color: silver;"><td>2</td><td>${row2.name}</td><td>${row2.time}</td></tr>\n`;
html += `<tr style="background-color: bronze;"><td>3</td><td>${row3.name}</td><td>${row3.time}</td></tr>\n`;
for (let i = 3; i < data.length; i++) {
  const row = data[i];
  html += `<tr><td>${i + 1}</td><td>${row.name}</td><td>${row.time}</td></tr>\n`;
}
html += '</table>';

// Write the CSS styles to a separate file
const fs = require('fs');
fs.writeFileSync('leaderboard.css', `
  table {
    border-collapse: collapse;
    width: 100%;
    background-image: linear-gradient(to bottom, #ffffff, #f2f2f2);
    margin: auto;
    justify-content: center;
    align-items: center;
  }

  th, td {
    text-align: center;
    padding: 8px;
  }

  th {
    background-color: #006400;
    color: white;
  }

  body {
    font-size: 20px;
  }
`);

// Link to the CSS file in your HTML file
html = `<html>\n<head>\n<link rel="stylesheet" href="leaderboard.css">\n</head>\n<body>\n${html}\n</body>\n</html>`;

// Write the HTML table to a file
fs.writeFileSync('leaderboard.html', html);

// Open the HTML file in the default browser
const { exec } = require('child_process');
exec(`open leaderboard.html`);
}
runCode(); // Run your code instantly the first time
// Close the current tab
setInterval(runCode, 6000); // Wait 60 seconds before running your code again
