const express = require("express");
const multer = require("multer");
// const { HL7Message } = require("hl7-standard");
// const HL7Message = require("hl7-standard");
const HL7 = require("hl7-standard");
const ExcelJS = require("exceljs");
const cors = require("cors");
const moment = require("moment");

const app = express();
const port = 3000;

// Configure multer for file upload
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(cors());

function convertDate(dateStr) {
  // Parse the date string using moment.js
  const parsedDate = moment(dateStr, "YYYYMMDD");
  // Format the parsed date as "Month Day, Year"
  const formattedDate = parsedDate.format("MMMM DD, YYYY");
  return formattedDate;
}

function convertDateTime(dateTimeStr) {
  // Parse the datetime string using moment.js
  const parsedDateTime = moment(dateTimeStr, "YYYYMMDDHHmmss");
  // Format the parsed datetime as "Month Day, Year hh:mm:ss"
  const formattedDateTime = parsedDateTime.format("MMMM DD, YYYY HH:mm:ss");
  return formattedDateTime;
}

// Handle POST request to /parseHL7
app.post("/parseHL7", upload.single("hl7File"), async (req, res) => {
  const hl7String = req.file.buffer.toString("utf-8");

  console.log("Received HL7 string:", hl7String); // Log received HL7 string
  const hl7 = new HL7(hl7String);
  hl7.transform();

  // Split the HL7 string into segments
  const segments = hl7String.split("\n");

  // Define arrays to store parsed data
  const patients = [];
  const messageHeaders = [];
  const notes = [];
  const obxData = [];

  let currentPatient = {}; // Initialize currentPatient object

  // Loop through each segment
  segments.forEach((segment) => {
    // Split the segment into fields
    const fields = segment.split("|");
    // Extract the segment type
    const segmentType = fields[0];

    if (segmentType === "PID") {
      // Initialize a new patient object for each PID segment
      currentPatient = {
        "Account #": fields[18],
        ID: fields[3],
        Sex: fields[8],
        Name: fields[5] ? fields[5].split("^").reverse().join(", ") : "",
        DOB: convertDate(fields[7]),
        // Initialize arrays to store patient-related data within the currentPatient object
        messageHeaders: [],
        notes: [],
        obxData: [],
      };
      patients.push(currentPatient); // Push the new patient object into the patients array
    }

    // if (segmentType === "MSH") {
    //   // Extract message header information
    //   const messageHeader = {
    //     App: fields[3],
    //     Facility: fields[4],
    //     "Msg Time": convertDate(fields[7]),
    //     "Control ID": fields[10],
    //     Type: fields[9].split("^")[0],
    //     Version: fields[12],
    //   };
    //   // currentPatient.messageHeaders.push(messageHeader); // Push message header data into patient's messageHeaders array
    //   if (!currentPatient.messageHeaders) {
    //     currentPatient.messageHeaders = []; // Initialize messageHeaders array if it's not already initialized
    //   }
    //   currentPatient.messageHeaders.push(messageHeader);
    // }

    if (segmentType === "NTE") {
      // Extract note
      const note = fields[3];
      // Push the note into the current patient's notes array
      currentPatient.notes.push(note);
    }

    if (segmentType === "OBX") {
      // Extract OBX data
      const obx = {
        Type: fields[3],
        Result: fields[5],
        Units: fields[6],
        Reference: fields[7],
        Abnormal: fields[8],
      };
      currentPatient.obxData.push(obx); // Push OBX data into patient's obxData array
    }
  });

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("HL7 Data");

    // Add columns to worksheet
    worksheet.columns = [
      { header: "Section", key: "section" },
      { header: "Property Name", key: "property" },
      { header: "Value", key: "value" },
    ];

    // Extract and add MSH segment data to worksheet
    const mshData = segments.find((segment) => segment.startsWith("MSH|"));
    if (mshData) {
      const mshFields = mshData.split("|");
      const mshHeaderData = [
        { section: "MSH Segment", property: "App", value: mshFields[3] },
        { section: "MSH Segment", property: "Facility", value: mshFields[4] },
        {
          section: "MSH Segment",
          property: "Msg Time",
          value: convertDateTime(mshFields[7]),
        },
        {
          section: "MSH Segment",
          property: "Control ID",
          value: mshFields[10],
        },
        {
          section: "MSH Segment",
          property: "Type",
          value: mshFields[9].split("^")[0],
        },
        { section: "MSH Segment", property: "Version", value: mshFields[12] },
      ];
      worksheet.addRows(mshHeaderData);
      // Add blank row after MSH segment data
      worksheet.addRow({});
    }

    // Add patient information and related data to worksheet
    patients.forEach((patient) => {
      // Add a blank row between patients
      if (patients.indexOf(patient) > 0) {
        worksheet.addRow({});
      }
      // Add patient information to worksheet
      const patientData = [
        {
          section: `Patient Information`,
          property: "Account #",
          value: patient["Account #"],
        },
        { section: `Patient Information`, property: "ID", value: patient.ID },
        { section: `Patient Information`, property: "Sex", value: patient.Sex },
        {
          section: `Patient Information`,
          property: "Name",
          value: patient.Name,
        },
        { section: `Patient Information`, property: "DOB", value: patient.DOB },
      ];
      worksheet.addRows(patientData);

      // Add message headers to worksheet
      patient.messageHeaders.forEach((header) => {
        const headerData = [
          { section: `Message Header`, property: "App", value: header.App },
          {
            section: `Message Header`,
            property: "Facility",
            value: header.Facility,
          },
          {
            section: `Message Header`,
            property: "Msg Time",
            value: header["Msg Time"],
          },
          {
            section: `Message Header`,
            property: "Control ID",
            value: header["Control ID"],
          },
          { section: `Message Header`, property: "Type", value: header.Type },
          {
            section: `Message Header`,
            property: "Version",
            value: header.Version,
          },
        ];
        worksheet.addRows(headerData);
      });

      patient.notes.forEach((note, index) => {
        worksheet.addRow({
          section: "Notes",
          property: `Note ${index + 1}`,
          value: note,
        });
      });

      // Add OBX data to worksheet
      patient.obxData.forEach((obx) => {
        // Extract the value for the label dynamically
        const labelComponents = obx.Type.split("^");
        const label = labelComponents.length >= 5 ? labelComponents[4] : "";

        // Create an array to store the OBX data for each property
        const obxData = [
          { section: `OBX Data`, property: label, value: obx.Result },
        ];

        // Add the OBX data array as a row in the worksheet
        worksheet.addRows(obxData);
      });
    });

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Set headers for response
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=data.xlsx");

    // Send Excel file as response
    res.send(buffer);
  } catch (error) {
    console.error("Error parsing HL7 message:", error);
    res.status(500).send("Error parsing HL7 message");
  }
});

app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});
