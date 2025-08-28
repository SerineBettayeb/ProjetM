// server.js
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.static("public")); // dossier pour HTML/JS

app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).send("Aucun fichier");

  const inputFile = req.file.path;
  const outputFile = path.join("uploads", `modifie_${req.file.originalname}`);

  try {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(inputFile);
    const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outputFile });

    for await (const worksheetReader of workbookReader) {
      const worksheetWriter = workbookWriter.addWorksheet(worksheetReader.name || "Feuille1");
      let headers = null;

      for await (const row of worksheetReader) {
        if (!headers) {
          headers = row.values.map(v => (v ? v.toString().trim() : ""));
          headers.push("Jour", "Mois", "Annee");
          worksheetWriter.addRow(headers).commit();
        } else {
          const rowValues = row.values.map(v => (v ? v.toString() : ""));
          const colIndex = headers.indexOf("DATE_CREATION");
          let jour = "", mois = "", annee = "";

          if (colIndex !== -1 && rowValues[colIndex]) {
            const dateValue = rowValues[colIndex];

            if (!isNaN(dateValue)) {
              const excelEpoch = new Date(Date.UTC(1899, 11, 30));
              const jsDate = new Date(excelEpoch.getTime() + dateValue*24*60*60*1000);
              jour = String(jsDate.getUTCDate()).padStart(2,"0");
              mois = String(jsDate.getUTCMonth()+1).padStart(2,"0");
              annee = String(jsDate.getUTCFullYear());
            } else if (typeof dateValue === "string") {
              const parts = dateValue.split("/");
              jour = parts[0]||"";
              mois = parts[1]||"";
              annee = parts[2]||"";
            }
          }

          rowValues.push(jour, mois, annee);
          worksheetWriter.addRow(rowValues).commit();
        }
      }
    }

    await workbookWriter.commit();
    res.download(outputFile);
  } catch (err) {
    console.error(err);
    res.status(500).send("Erreur lors du traitement du fichier");
  }
});

app.listen(3000, () => console.log("🌐 Serveur lancé sur http://localhost:3000"));
