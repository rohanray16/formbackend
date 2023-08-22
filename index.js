const express = require("express");
const app = express();
const bodyParser = require("body-parser");
const path = require("path");
const cors = require("cors");
const fs = require("fs");
const exceljs = require("exceljs");

const PORT = 3001;
const EXCEL_FILE = "form.xlsx";

app.use(cors());
app.use(bodyParser.json());



app.get("/download-excel", (req, res) => {
  res.download(EXCEL_FILE, "form.xlsx", (err) => {
    if (err) {
      res.status(500).json({ error: "Error downloading file." });
    }
  });
});

// Serve the static build folder (if you have a React build)
// app.use(express.static(path.join(__dirname, "build")));

app.post("/submit-form", (req, res) => {
  const formData = req.body;

  // Create or open the Excel workbook
  let workbook;
  if (fs.existsSync(EXCEL_FILE)) {
    workbook = new exceljs.Workbook();
    workbook.xlsx.readFile(EXCEL_FILE).then(() => {
      processFormData(workbook, formData);
    });
  } else {
    workbook = new exceljs.Workbook();
    processFormData(workbook, formData);
  }

  res.status(200).json({ message: "Form data submitted successfully." });
});
function getMaxSlNo(sheet) {
  let maxSlNo = 0;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Skip the header row
      const slNo = row.getCell(1).value;
      if (slNo > maxSlNo) {
        maxSlNo = slNo;
      }
    }
  });
  return maxSlNo;
}
function processFormData(workbook, formData) {
  let sheet = workbook.getWorksheet("Form Data");
  if (!sheet) sheet = workbook.addWorksheet("Form Data");
  sheet.columns = [
    { header: "SL No", key: "slNo", width: 10 },
    { header: "Name", key: "name", width: 20 },
    { header: "Connected to", key: "connectedTo", width: 20 },
    { header: "Category", key: "category", width: 15 },
    { header: "Relative's Name", key: "relativeName", width: 20 },
    { header: "Relative's Age", key: "relativeAge", width: 15 },
    { header: "Relative's Relation", key: "relativeRelation", width: 20 },
    { header: "Payment Date", key: "paymentDate", width: 15 },
    { header: "Payment Amount", key: "paymentAmount", width: 15 },
    { header: "Any Concerns", key: "concerns", width: 30 },
  ];

  // Add the form data to the Excel sheet
  const maxSlNo = getMaxSlNo(sheet);
  const slNo = maxSlNo + 1;
  const pd = formData.paymentDate
    ? ("" + formData.paymentDate).split("T")[0].split("-")
    : null;
  const paymentDate = pd ? `${pd[2]}/${pd[1]}/${pd[0]}` : "";
  const rowData = {
    slNo,
    name: formData.name,
    connectedTo: formData.connectedTo,
    category: formData.category,
    relativeName: "",
    relativeAge: "",
    relativeRelation: "",
    paymentDate,
    paymentAmount: formData.paymentAmount,
    concerns: formData.concerns,
  };

  sheet.addRow(rowData);

  // Add relative data if available
  if (formData.relatives.length > 0) {
    formData.relatives.forEach((relative) => {
      const relativeRowData = {
        slNo: "",
        name: "",
        connectedTo: "",
        category: "",
        relativeName: relative.name,
        relativeAge: relative.age,
        relativeRelation: relative.relation,
        paymentDate: "",
        paymentAmount: "",
        concerns: "",
      };

      sheet.addRow(relativeRowData);
    });
  }

  // Save the workbook
  workbook.xlsx.writeFile(EXCEL_FILE);
}

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
