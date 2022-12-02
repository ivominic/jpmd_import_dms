const fs = require("fs");
const readXlsxFile = require("read-excel-file/node");

const excelFilePath = "C:/Users/PC/Desktop/contracts_use_md_2019.xlsx";
const sourceFolder = "C:/Users/PC/Desktop/md/ugovori_source/";
const destinationFolder = "C:/Users/PC/Desktop/md/ugovori_destination/";

readXlsxFile(excelFilePath).then((rows) => {
  rows.forEach((documentRow) => {
    if (documentRow[10] === "1" && documentRow[19]?.includes("lokacija")) {
      let subfolder = documentRow[13];
      let newFilename = formFileName(documentRow);
      if (newFilename) {
        moveFiles(sourceFolder + subfolder + "/1.pdf", destinationFolder + newFilename);
      }
      //console.log("******    ", newFilename);
    }
  });

  // `rows` is an array of rows
  // each row being an array of cells.
});

function formFileName(row) {
  let retVal = "";
  let contractDate = row[5];
  let contractNumber = row[20]
    .replaceAll(" ", "")
    .substring(0, row[20].length - 7)
    .replaceAll("/", "_");
  let splitArray = row[19].split("lokacija");
  let ownerName = splitArray[0].trim().replaceAll('"', "");
  let secondPart = splitArray[1].trim();
  let locationNumber = "";
  let municipality = "";
  if (secondPart.includes(" ")) {
    let secondArray = secondPart.split(" ");
    locationNumber = secondArray[0].trim();
    municipality = secondArray[1].trim();
    municipality === "Herceg" && (municipality = "Herceg Novi");
  } else {
    if (secondPart.startsWith("D") || secondPart.startsWith("L")) {
      municipality = "Ulcinj";
      locationNumber = secondPart;
    }
  }

  let blnMunicipality = ["Bar", "Budva", "Herceg Novi", "Kotor", "Tivat", "Ulcinj"].includes(municipality);

  if (locationNumber && blnMunicipality) {
    retVal = `${municipality}#${locationNumber}#${contractNumber}#${contractDate}#${ownerName}.pdf`;
  } else {
    console.log("-----", `${municipality}#${locationNumber}#${contractNumber}#${contractDate}#${ownerName}.pdf`);
  }
  return retVal;
}

function moveFiles(sourceFile, destinationFile) {
  if (fs.existsSync(destinationFile)) {
    console.log("Postoji fajl", destinationFile);
  } else {
    fs.rename(sourceFile, destinationFile, function (err) {
      if (err) console.error(sourceFile, err);
    });
  }
}
