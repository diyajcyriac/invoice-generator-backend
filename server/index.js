const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Function to parse different weight units and convert them to kilograms
const parseGms = (gms) => {
  if (typeof gms === 'string') {
    gms = gms.toLowerCase().trim();
    if (gms.endsWith('kg')) {
      return parseFloat(gms.replace(' kg', '').replace('kg', ''));
    } else if (gms.endsWith('g')) {
      return parseFloat(gms.replace(' g', '').replace('g', '')) / 1000;
    } else if (gms.endsWith('lb')) {
      return parseFloat(gms.replace(' lb', '').replace('lb', '')) / 2.20462;
    }
  }
  return parseFloat(gms) || 0;
};

// Function to calculate optimal adjusted rates
function calculateOptimalAdjustedRates(rates, kilograms) {
  const adjustments = [0.00, -0.01, 0.01, 0.02, -0.02, 0.03, -0.03];

  // Function to round to two decimal places
  function roundToTwoDecimalPlaces(value) {
    return Math.round((value + Number.EPSILON) * 100) / 100;
  }

  const adjustedRates = rates.map(rate => roundToTwoDecimalPlaces(rate));

  const minKilogramIndices = [];
  const sortedKilograms = [...kilograms].sort((a, b) => a - b).slice(0, 10);
  const usedIndices = new Set();

  for (let minKg of sortedKilograms) {
    let index = kilograms.findIndex((kg, i) => kg === minKg && !usedIndices.has(i));
    usedIndices.add(index);
    minKilogramIndices.push(index);
  }

  let sumRateKg = 0;
  let sumAdjustedRateKg = 0;

  for (let i = 0; i < kilograms.length; i++) {
    if (!minKilogramIndices.includes(i)) {
      sumRateKg += rates[i] * kilograms[i];
      sumAdjustedRateKg += adjustedRates[i] * kilograms[i];
    }
  }

  const c = sumRateKg - sumAdjustedRateKg;

  let d = 0;
  for (let idx of minKilogramIndices) {
    d += rates[idx] * kilograms[idx];
  }

  const b = c + d;

  const x = [];
  const y = [];
  for (let i = 0; i < 10; i++) {
    x[i] = adjustedRates[minKilogramIndices[i]];
    y[i] = kilograms[minKilogramIndices[i]];
  }

  let closestDiff = Infinity;
  let bestX = adjustedRates.slice();

  for (let adj1 of adjustments) {
    for (let adj2 of adjustments) {
      for (let adj3 of adjustments) {
        for (let adj4 of adjustments) {
          for (let adj5 of adjustments) {
            for (let adj6 of adjustments) {
              for (let adj7 of adjustments) {
                for (let adj8 of adjustments) {
                  for (let adj9 of adjustments) {
                    for (let adj10 of adjustments) {
                      const newX = [
                        roundToTwoDecimalPlaces(x[0] + adj1),
                        roundToTwoDecimalPlaces(x[1] + adj2),
                        roundToTwoDecimalPlaces(x[2] + adj3),
                        roundToTwoDecimalPlaces(x[3] + adj4),
                        roundToTwoDecimalPlaces(x[4] + adj5),
                        roundToTwoDecimalPlaces(x[5] + adj6),
                        roundToTwoDecimalPlaces(x[6] + adj7),
                        roundToTwoDecimalPlaces(x[7] + adj8),
                        roundToTwoDecimalPlaces(x[8] + adj9),
                        roundToTwoDecimalPlaces(x[9] + adj10)
                      ];

                      let newC = 0;
                      for (let i = 0; i < x.length; i++) {
                        newC += newX[i] * y[i];
                      }

                      const diff = Math.abs(newC - b);

                      if (diff < 0.01 && diff < closestDiff) {
                        closestDiff = diff;
                        bestX = newX.slice();
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  const newAdjustedRates = adjustedRates.slice();
  for (let i = 0; i < 10; i++) {
    newAdjustedRates[minKilogramIndices[i]] = roundToTwoDecimalPlaces(bestX[i]);
  }

  calculateOptimalAdjustedRates.closestDiff = closestDiff;

  return newAdjustedRates.map(rate => roundToTwoDecimalPlaces(rate));
}

// Route to handle file upload and processing
app.post('/upload', upload.single('file'), (req, res) => {
  const file = req.file;
  const additionalFreight = parseFloat(req.body.additionalFreight);

  if (!file || isNaN(additionalFreight)) {
    return res.status(400).json({ error: 'File and additional freight are required' });
  }

  const workbook = XLSX.readFile(file.path);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (!Array.isArray(data) || data.length === 0) {
    return res.status(400).json({ error: 'Invalid data format in the Excel sheet' });
  }

  // Adding headers to the new columns
  data[0].push(
    'Total no. of kgs', 
    'Rates with freight added', 
    'Price with added freight included', 
    'Rates per kg', 
    'Adjusted rates per kg',
    'Price with adjusted Rates'
  );

  const totalCartons = data.slice(1).reduce((sum, row) => sum + parseFloat(row[2]), 0);
  const c = additionalFreight / totalCartons;

  const ratesPerKg = [];
  const kilograms = [];

  // Processing each row to calculate new values
  for (let i = 1; i < data.length; i++) {
    if (data[i].length < 8) {
      continue;
    }

    const units = parseFloat(data[i][6]);
    const gms = parseGms(data[i][5]);
    const totalKgs = units * gms;
    data[i].push(totalKgs);

    const noOfPackets = parseFloat(data[i][3]);
    const ratePerUnit = parseFloat(data[i][7]);
    const ratesWithFreightAdded = (c / noOfPackets) + ratePerUnit;
    const priceWithFreightIncluded = units * ratesWithFreightAdded;
    const ratePerKg = priceWithFreightIncluded / totalKgs;

    data[i].push(ratesWithFreightAdded, priceWithFreightIncluded, ratePerKg);

    ratesPerKg.push(ratePerKg);
    kilograms.push(totalKgs);
  }

  const adjustedRatesPerKg = calculateOptimalAdjustedRates(ratesPerKg, kilograms);

  // Adding adjusted rates to the data
  for (let i = 1; i < data.length; i++) {
    if (data[i].length < 13) {
      continue;
    }

    data[i].push(adjustedRatesPerKg[i - 1]);
  }
  
  // Adding final price with adjusted rates
  for (let i = 1; i < data.length; i++) {
    if (data[i].length < 14) {
      continue;
    }
    const priceWithAdjustedRate = data[i][12];
    data[i].push(priceWithAdjustedRate);
  }

  const newSheet = XLSX.utils.aoa_to_sheet(data);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);

  // Define the downloads directory
  const downloadsDir = path.join(__dirname, '..', 'downloads');
  if (!fs.existsSync(downloadsDir)) {
    fs.mkdirSync(downloadsDir);
  }

  // Save the output file
  const outputFilePath = path.join(downloadsDir, 'output.xlsx');
  XLSX.writeFile(newWorkbook, outputFilePath);

  // Prepare the invoice file with specific columns
  const invoiceData = data.map((row, index) => {
    if (index === 0) {
      return [
        'Sl.No.',
        'Product',
        'HSN No.',
        'Cartons',
        'no. of packets',
        '',
        'gms',
        'Total no. of kgs',
        'Adjusted rates per kg',
        'Price with adjusted Rates'
      ];
    }
    return [
      row[0], // Sl.No.
      row[1], // Product
      row[9], // HSN No.
      row[2], // Cartons
      row[3], // no. of packets
      row[4],     // <blank>
      row[5], // gms
      row[8], // Total no. of kgs
      row[14],// Adjusted rates per kg
      row[15] // Price with adjusted Rates
    ];
  });

  const invoiceSheet = XLSX.utils.aoa_to_sheet(invoiceData);
  const invoiceWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(invoiceWorkbook, invoiceSheet, 'Invoice');

  const invoiceFilePath = path.join(downloadsDir, 'invoice.xlsx');
  XLSX.writeFile(invoiceWorkbook, invoiceFilePath);

  // Send response with download links
  res.json({ outputFile: `https://invoice-generator-backend-uock.onrender.com/downloads/output.xlsx`, invoiceFile: `https://invoice-generator-backend-uock.onrender.com/downloads/invoice.xlsx` });
});

// Serve files in the downloads directory
app.use('/downloads', express.static(path.join(__dirname, '..', 'downloads')));

const port = process.env.PORT || 5000;;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
