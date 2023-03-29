function runCode() {
// Import the xlsx library
console.log('Importing xlsx library...');
console.log(new Date().toLocaleTimeString());
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

// Convert time in mm:ss format to seconds
console.log('Converting time in mm:ss format to seconds...');
data.forEach((row) => {
  row.time = convertToSeconds(row.time);
});
console.log('Converted time in mm:ss format to seconds.');

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
}
function convertToSeconds(time) {
  if (typeof time === 'number') {
    return time;
  }

  if (typeof time !== 'string' || !time.includes(':')) {
    throw new Error('Invalid time format');
  }

  const [minutes, seconds] = time.split(':');
  return Number(minutes) * 60 + Number(seconds);
}
//defines html elements
let html = '<div style="display: flex; justify-content: center; align-items: center">';
html += '<div style="text-align:center; width: 60%; max-width: 600px;">';
html += '<h1 style="font-family: sans-serif; color: #fff; margin-bottom: 20px; background-color: #4da6ff; padding: 5px; border-radius: 5px;">Pirbright Glove Trail</h1>';
html += '<img src="mos.jpg" alt="mos" style="width: 10%; height: auto; margin: 0 auto; align-items: center;"></img>'
html += '<table style="border-collapse: collapse; width: 100%; margin-top: 50px; background-color: #0077b3; border-radius: 10px; overflow: hidden;">\n';
html += '<tr style="background-color: gold;"><th style="padding: 10px; border-bottom: 2px solid #333;">Rank</th><th style="padding: 10px; border-bottom: 2px solid #333;">Name</th><th style="padding: 10px; border-bottom: 2px solid #333;">Time</th></tr>\n';
//first row (first position)
const row1 = data[0];
if (row1) {
    html += `<tr style="background-color: gold;"><td style="padding: 10px; border-bottom: 1px solid #333;">1</td><td style="padding: 10px; border-bottom: 1px solid #333;">${row1.name}</td><td style="padding: 10px; border-bottom: 1px solid #333;">${row1.time}</td></tr>\n`;
}
//second row (second position)
const row2 = data[1];
if (row2) {
    html += `<tr style="background-color: #C0C0C0;"><td style="padding: 10px;">2</td><td style="padding-left:.5rem;padding-right:.5rem;border-bottom:.0625rem solid rgba(0,0,0,.05);">${row2.name}</td><td style="padding-left:.5rem;padding-right:.5rem;border-bottom:.0625rem solid rgba(0,0,0,.05);">${row2.time}</td></tr>\n`;
}
//third row (third position)
const row3 = data[2];
if (row3) {
    html += `<tr style="background-color: #CD7F32;"><td style="padding: 10px;">3</td><td style="padding-left:.5rem;padding-right:.5rem;">${row3.name}</td><td style="padding-left:.5rem;padding-right:.5rem;">${row3.time}</td></tr>\n`;
}
//all other rows
for (let i = 3; i < data.length; i++) {
    const row = data[i];
    html += `<tr><td style="padding-left:.5rem;padding-right:.5rem; border-top: 1px solid black";>${i + 1}</td><td style="padding-left:.5rem;padding-right:.5rem; border-top: 1px solid black;">${row.name}</td><td style="padding-left:.5rem;padding-right:.5rem; border-top: 1px solid black;">${row.time}</td></tr>\n`;
}
html += '</table>';
html += '<img src="r.jpg" alt="mos" style="width: 30%; height: auto; margin: 0 auto; align-items: center;"></img>'
html += '</div>';
html += '</div>';


// Write the CSS styles to a separate file
const fs = require('fs');
fs.writeFileSync('leaderboard.css', `
  table {
    border-collapse: collapse;
    width: 100%;
    background-image: linear-gradient(to bottom, #cd9932, #FF7276);
    margin: auto;
    justify-content: center;
    align-items: center;
    border-collapse: separate;
    border-spacing: 10px;
    
  }
  body {

  }
  h1 {
    font-family: Arial, sans-serif;
    font-size: 35px;
    border-radius: 1em;
    color: white;
    padding: 8px;
    text font-weight: bold;

  }
  th, td {
    font-family: Arial, sans-serif;
  }

  th, td {
    text-align: center;
    padding: 8px;
  }

  th {
    background-color: #00008B;
    color: white;
  }

  th {
    font-size: 30px;
    text font-weight: bold;
  }
  
  td {
    border-radius: 1em;
  }
  th {
    border-radius: 3em;
  }

  tr {
    opacity: 0.7;
    color: light grey;

  }
  
}
`);

// Link to the CSS file in your HTML file
html = `<html>\n<head>\n<link rel="stylesheet" href="leaderboard.css">\n</head>\n<body>\n${html}\n</body>\n</html>`;

// Write the HTML table to a file
fs.writeFileSync('leaderboard.html', html);

// Open the HTML file in the default browser
const { exec } = require('child_process');
exec(`open leaderboard.html`);
console.log('Opening site.')
console.log(new Date().toLocaleTimeString());

runCode(); // Run your code instantly the first time
// Close the current tab
setInterval(runCode, 10000); // Wait 10 seconds before running your code again
