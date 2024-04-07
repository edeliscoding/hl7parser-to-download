const express = require("express");
const multer = require("multer");
// const { HL7Message } = require("hl7-standard");
// const HL7Message = require("hl7-standard");
const HL7 = require("hl7-standard");
const ExcelJS = require("exceljs");
const cors = require("cors");

const app = express();
const port = 3000;

// Configure multer for file upload
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(cors());

// Handle POST request to /parseHL7
app.post("/parseHL7", upload.single("hl7File"), async (req, res) => {
  const hl7String = req.file.buffer.toString("utf-8");

  console.log("Received HL7 string:", hl7String); // Log received HL7 string
  const hl7 = new HL7(hl7String);
  hl7.transform();

  try {
    // Get values using the get method
    const patientInformation = {
      "Account #": hl7.get("PID.18"),
      // ID: hl7.get("PID.3") + ", " + hl7.get("PID.20"),
      ID: hl7.get("PID.3"),
      // "Second ID": hl7.get("PID.20"),
      Sex: hl7.get("PID.8"),
      Name: hl7.get("PID.5.2") + ", " + hl7.get("PID.5.1"),
      DOB: hl7.get("PID.7"),
      // Address information is not provided in the given HL7 message
    };

    // Get values using the get method
    const messageHeader = {
      App: hl7.get("MSH.3"),
      Facility: hl7.get("MSH.4"),
      "Msg Time": hl7.get("MSH.7"),
      "Control ID": hl7.get("MSH.10"),
      Type: hl7.get("MSH.9.1"),
      Version: hl7.get("MSH.12"),
    };

    const notes = [];
    const obrSegments = hl7.getSegments("OBR");
    for (const obr of obrSegments) {
      const nteGroup = hl7.getSegmentsAfter(obr, "NTE", true, ["OBR", "ORC"]);
      const notesData = nteGroup.map((segment) => ({
        Note: segment.get("NTE.3"),
      }));
      notes.push({ Notes: notesData });
    }

    const obxData = [];
    const obxSegments = hl7.getSegments("OBX");
    for (const obx of obxSegments) {
      obxData.push({
        Type: obx.get("OBX.3"),
        Result: obx.get("OBX.5"),
        Units: obx.get("OBX.6"),
        Reference: obx.get("OBX.7"),
        Abnormal: obx.get("OBX.8"),
      });
    }

    // const formattedOutput = [
    //   { "Message Header": messageHeader },
    //   { "Patient Information": patientInformation },
    //   ...notes,
    //   { "OBX Data": obxData },
    // ];

    // res.json(formattedOutput);
    // Create workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("HL7 Data");

    // Add data to worksheet
    worksheet.columns = [
      { header: "Section", key: "section" },
      { header: "Property Name", key: "property" },
      { header: "Value", key: "value" },
    ];
    const data = [
      {
        section: "Patient Information",
        property: "Account #",
        value: patientInformation["Account #"],
      },
      {
        section: "Patient Information",
        property: "ID",
        value: patientInformation.ID,
      },
      {
        section: "Patient Information",
        property: "Sex",
        value: patientInformation.Sex,
      },
      {
        section: "Patient Information",
        property: "Name",
        value: patientInformation.Name,
      },
      {
        section: "Patient Information",
        property: "DOB",
        value: patientInformation.DOB,
      },
      { section: "Message Header", property: "App", value: messageHeader.App },
      {
        section: "Message Header",
        property: "Facility",
        value: messageHeader.Facility,
      },
      {
        section: "Message Header",
        property: "Msg Time",
        value: messageHeader["Msg Time"],
      },
      {
        section: "Message Header",
        property: "Control ID",
        value: messageHeader["Control ID"],
      },
      {
        section: "Message Header",
        property: "Type",
        value: messageHeader.Type,
      },
      {
        section: "Message Header",
        property: "Version",
        value: messageHeader.Version,
      },
      ...notes.flatMap((note, index) =>
        note.Notes.map((noteData) => ({
          section: "Notes",
          property: `Note ${index + 1}`,
          value: noteData.Note,
        }))
      ),
      ...obxData.map((obx, index) => ({
        section: "OBX Data",
        property: `Type ${index + 1}`,
        value: obx.Type,
      })),
      ...obxData.map((obx, index) => ({
        section: "OBX Data",
        property: `Result ${index + 1}`,
        value: obx.Result,
      })),
      ...obxData.map((obx, index) => ({
        section: "OBX Data",
        property: `Units ${index + 1}`,
        value: obx.Units,
      })),
      ...obxData.map((obx, index) => ({
        section: "OBX Data",
        property: `Reference ${index + 1}`,
        value: obx.Reference,
      })),
      ...obxData.map((obx, index) => ({
        section: "OBX Data",
        property: `Abnormal ${index + 1}`,
        value: obx.Abnormal,
      })),
    ];
    worksheet.addRows(data);

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
    onsole.error("Error parsing HL7 message:", error);
  }
});
app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});
