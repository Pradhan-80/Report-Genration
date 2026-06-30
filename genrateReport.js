
const fs = require("fs");
const XLSX = require("xlsx");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const express = require("express");
const multer = require("multer");
const ImageModule = require("docxtemplater-image-module-free");
const PORT = process.env.PORT || 3000;

const sizeOf = require("image-size");

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));

const upload = multer({
  dest: "uploads/",
});

//const EXCEL_FILE = "input.xlsx";
const DOCX_TEMPLATE = "Report_GGNIN.docx";
const DOCX_OUTPUT = "output.docx";

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/index.html");
});


function readSheet(wb, name) {
  const sheet = wb.Sheets[name];

  if (!sheet) return [];

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
  });
}

function findHeader(rows, keyword) {
  return rows.findIndex((r) =>
    r.some((c) => String(c).toLowerCase().includes(keyword.toLowerCase())),
  );
}


function parseDetails(rows) {
  const details = {};

  rows.forEach((r) => {
    const key = String(r[0] || "").trim();
    const value = r[1] || "";

    if (!key) return;

    details[key] = value;
  });

  return {
    receivedDate: details["Received Date"] || "",
    cardsReceived: details["No. of Cards Received in LAB"] || "",
    cardType: details["Card_Type"] || "",
    circle: details["Circle"] || "",
    cardProfile: details["Card Profile"] || "",
    vendor: details["Vendor"] || "",
    lotSize: details["Lot Size"] || "",
    poNo: details["PO No."] || "",
    invoiceNo: details["Invoice No."] || "",
    startingICCID:
      details["Starting ICCID of Lot received in WH"] || "",
    endingICCID:
      details["Ending ICCID of Lot received in WH"] || "",
    lotReceivedDate:
      details["Lot Received Date in WH"] || "",
    testingCompleted:
      details["Testing Completed"] || "",
    requestedBy:
      details["Requested By"] || "",
  };
}

function parseThickness(rows) {
  const i = findHeader(rows, "iccid");

  if (i === -1) return [];

  return rows
    .slice(i + 1)
    .filter((r) => r[1])
    .map((r) => ({
      iccid: r[0],
      specification: r[1],
      observed: r[2],
      remarks: r[3],
    }));
}

function parseGoNoGo(rows) {
  const i = findHeader(rows, "parameter");

  if (i === -1) return [];

  let iccid = null;
  const out = [];

  rows.slice(i + 1).forEach((r) => {
    if (r[1]) iccid = r[1];

    if (!r[2]) return;

    out.push({
      iccid,
      parameter: r[1],
      specification: r[2],
      passed: r[3],
    });
  });

  return out;
}

function parseVoltage(rows) {
  const i = findHeader(rows, "class a");

  if (i === -1) return [];

  return rows
    .slice(i + 1)
    .filter((r) => r[1])
    .map((r) => ({
      iccid: r[0],
      classA: r[1],
      classB: r[2],
      classC: r[3],
      maxV: r[4],
      minV: r[5],
      remarks: r[6],
    }));
}

function parseOnline(rows) {
  const i = findHeader(rows, "permanent imsi");

  if (i === -1) return [];

  return rows
    .slice(i + 1)
    .filter((r) => r[1])
    .map((r) => ({
      iccid: r[0],
      permanentImsi: r[1],
      msisdn: r[2],
      dataCheck: r[3],
      incomingCall: r[4],
      outgoingCall: r[5],
      incomingSms: r[6],
      outgoingSms: r[7],
      handset: r[8],
    }));
}

function parseDSA(rows) {
  const i = findHeader(rows, "ICCID");

  if (i === -1) return [];

  return rows
    .slice(i + 1)
    .filter((r) => r[1])
    .map((r) => ({
      iccid: r[1],
      tempImsi: r[2],
      permImsi: r[3],
      msisdn: r[4],
      smsp: r[5],
      incoming: r[6],
      outgoing: r[7],
      handset: r[8],
      circle: r[9],
    }));
}

function flattenDSTK(dstkObj) {
  const rows = [];

  Object.entries(dstkObj).forEach(([iccid, tests]) => {
    tests.forEach((t) => {
      rows.push({
        iccid,
        testCaseId: t.testCaseId,
        summary: t.summary,
        description: t.description,
        result: t.result,
        remarks: t.remarks,
      });
    });
  });

  return rows;
}

function parseDSTK(rows) {
  const i = findHeader(rows, "Test Case ID");

  if (i === -1) return {};

  let currentICCID = null;

  const grouped = {};

  rows.slice(i + 1).forEach((r) => {
    if (r[0]) currentICCID = r[0];

    if (!r[1]) return;

    if (!grouped[currentICCID]) {
      grouped[currentICCID] = [];
    }

    grouped[currentICCID].push({
      testCaseId: r[1],
      summary: r[2],
      description: r[3],
      result: r[4],
      remarks: r[5],
    });
  });

  return grouped;
}

function parseOta(rows) {
  return rows
    .filter((r) => r[0]?.startsWith("#"))
    .map((r) => ({
      log: r[0],
      status: r[1] || "NA",
    }));
}


function parseExcel(file) {
  const wb = XLSX.readFile(file);

  return {
    Details: parseDetails(readSheet(wb, "Details")),
    Online: parseOnline(readSheet(wb, "Online")),

    thickness: parseThickness(readSheet(wb, "Thickness")),

    voltage: parseVoltage(readSheet(wb, "Voltage")),

    goNoGo: parseGoNoGo(readSheet(wb, "GnG")),

    DSA: parseDSA(readSheet(wb, "DSA")),

    dstk: flattenDSTK(parseDSTK(readSheet(wb, "DSTK"))),

    ota: parseOta(readSheet(wb, "OTA")),
  };
}


function suppressRepeatedValues(rows, key) {
  let lastValue = null;

  return rows.map((row) => {
    if (row[key] === lastValue) {
      return {
        ...row,
        [key]: "",
      };
    }

    lastValue = row[key];

    return row;
  });
}


function generateDocx(data, workOrder, simPic1, simPic2) {
  // optional cleanup
  data.goNoGo = suppressRepeatedValues(data.goNoGo, "iccid");
  data.dstk = suppressRepeatedValues(data.dstk, "iccid");

  const content = fs.readFileSync(DOCX_TEMPLATE, "binary");

  const zip = new PizZip(content);
  const imageModule = new ImageModule({
    getImage(path) {
      return fs.readFileSync(path);
    },

    getSize(path) {
      return [600, 350];
    },
  });

 
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [imageModule],
  });

 
  doc.render({
    workOrder,

    simPic1: simPic1?.path,

    simPic2: simPic2?.path,

    ...data,
  });

  const buffer = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  // fs.writeFileSync(DOCX_OUTPUT, buffer);
  return buffer;
}

// app.post("/generate", upload.single("excelFile"), (req, res) => {
app.post(
  "/generate",
  upload.fields([
    {
      name: "excelFile",
      maxCount: 1,
    },
    {
      name: "simPic1",
      maxCount: 1,
    },
    {
      name: "simPic2",
      maxCount: 1,
    },
  ]),
  (req, res) => {
    try {
      const workOrder = req.body.workOrder;
      //  const excelFilePath = req.file.path;
      const excelFilePath = req.files.excelFile[0].path;
      
      const simPic1 = req.files.simPic1?.[0];
      const simPic2 = req.files.simPic2?.[0];

      const data = parseExcel(excelFilePath);

      // generateDocx(data, workOrder);

      // console.log("SUCCESS → output.docx generated");

      // res.download(DOCX_OUTPUT);
      const buffer = generateDocx(data, workOrder, simPic1, simPic2);

      res.setHeader(
        "Content-Disposition",
        `attachment; filename=Report_${workOrder}.docx`,
      );

      res.send(buffer);
    } catch (err) {
      console.error(err);

      res.status(500).send("Failed to generate DOCX");
    }
  },
);

// app.listen(3000, () => {
//   console.log("Server running at http://localhost:3000");
// });
app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`);
});
