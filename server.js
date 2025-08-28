const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const upload = multer({ dest: "/tmp" });

app.use(express.static("public"));

// Stockage temporaire des fichiers traités
const processedFiles = {};

// Route upload
app.post("/upload", upload.single("file"), async (req, res) => {
  console.log("📥 Requête reçue sur /upload");

  if (!req.file) {
    console.warn("⚠️ Aucun fichier reçu");
    return res.status(400).send("Aucun fichier");
  }

  const inputFile = req.file.path;
  const outputFile = path.join("/tmp", `modifie_${req.file.originalname}`);

  console.log(`📁 Fichier reçu: ${req.file.originalname}`);
  console.log("🛠 Traitement en arrière-plan...");

  // Lance le traitement asynchrone
  (async () => {
    try {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(inputFile);
      const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outputFile });

      for await (const worksheetReader of workbookReader) {
        console.log(`📄 Lecture de la feuille: ${worksheetReader.name || "Feuille1"}`);
        const worksheetWriter = workbookWriter.addWorksheet(worksheetReader.name || "Feuille1");

        let headers = null;
        let batch = [];

        for await (const row of worksheetReader) {
          if (!headers) {
            headers = row.values.map(v => (v ? v.toString().trim() : ""));
            headers.push("Jour", "Mois", "Annee");
            worksheetWriter.addRow(headers).commit();
            console.log(`📝 Headers ajoutés: ${headers.join(", ")}`);
          } else {
            const rowValues = row.values.map(v => (v ? v.toString() : ""));
            const colIndex = headers.indexOf("DATE_CREATION");
            let jour = "", mois = "", annee = "";

            if (colIndex !== -1 && rowValues[colIndex]) {
              const dateValue = rowValues[colIndex];
              if (!isNaN(dateValue)) {
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const jsDate = new Date(excelEpoch.getTime() + dateValue * 24 * 60 * 60 * 1000);
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
            batch.push(rowValues);

            if(batch.length >= 1000){
              worksheetWriter.addRows(batch);
              batch = [];
            }
          }
        }

        if(batch.length > 0){
          worksheetWriter.addRows(batch);
        }
      }

      await workbookWriter.commit();
      console.log(`✅ Fichier traité: ${outputFile}`);

      processedFiles[req.file.originalname] = outputFile;

    } catch(err) {
      console.error("❌ Erreur pendant le traitement:", err);
    }
  })();

  // Réponse immédiate
  res.send(`
    Fichier reçu. Traitement en cours.<br>
    Vérifiez le statut ici: <a href="/status/${req.file.originalname}">/status/${req.file.originalname}</a>
  `);
});

// Statut du fichier
app.get("/status/:filename", (req, res) => {
  const filePath = processedFiles[req.params.filename];
  if(filePath && fs.existsSync(filePath)){
    res.json({ ready: true, url: `/download/${req.params.filename}` });
  } else {
    res.json({ ready: false });
  }
});

// Téléchargement
app.get("/download/:filename", (req, res) => {
  const filePath = processedFiles[req.params.filename];
  if(filePath && fs.existsSync(filePath)){
    res.download(filePath, `modifie_${req.params.filename}`);
  } else {
    res.status(404).send("Fichier non disponible ou traitement pas encore terminé");
  }
});

// Render impose process.env.PORT
const PORT = process.env.PORT;
if (!PORT) {
  console.error("❌ PORT non défini !");
  process.exit(1);
}

app.listen(PORT, () => console.log(`🌐 Serveur lancé sur le port ${PORT}`));
