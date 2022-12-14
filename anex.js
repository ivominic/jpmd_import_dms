const fs = require("fs");
const readXlsxFile = require("read-excel-file/node");

const excelDFilePath = "C:/Users/PC/Desktop/aneksi ugovora koriscenje md/anex_d.xlsx";
const excelKFilePath = "C:/Users/PC/Desktop/aneksi ugovora koriscenje md/anex_k.xlsx";
const sourceFolder = "C:/Users/PC/Desktop/aneksi ugovora koriscenje md/anex_source/";
const destinationFolder = "C:/Users/PC/Desktop/aneksi ugovora koriscenje md/anex_destination/";

let dRows, kRows;

readXlsxFile(excelDFilePath)
  .then((rows) => {
    dRows = rows;
  })
  .then(
    readXlsxFile(excelKFilePath).then((rows) => {
      kRows = rows;
      return kRows;
    })
  )
  .then((r) => {
    console.log(dRows.length, kRows.length);
    dRows.forEach((documentRow) => {
      let retArray = readKFile(documentRow[20].trim().replaceAll(" "));
      if (retArray.length) {
        let subfolder = documentRow[13];
        let newFilename = formFileName(retArray);
        if (newFilename) {
          moveFiles(sourceFolder + subfolder + "/1.pdf", destinationFolder + newFilename);
        }
      }
    });
  });

function readKFile(dmsName) {
  let retVal = [];
  kRows.forEach((documentRow) => {
    let docName = documentRow[7]?.toLowerCase().split(" od ");
    if (!retVal.length && docName?.length === 2) {
      let name = docName[0].trim();
      let date = docName[1].trim();
      let year = date.substring(date.length - 4, date.length);
      if (date.endsWith(".")) {
        year = date.substring(date.length - 5, date.length - 1);
      } else {
        date = date + ".";
      }
      let municipality = documentRow[0]?.trim();
      let location = documentRow[1]?.trim().replaceAll(" ");
      if (location?.endsWith(".")) {
        location = location?.substring(0, location?.length - 1);
      }
      let user = documentRow[3]?.trim();
      if (dmsName === `${name}-(${year})`) {
        retVal.push(municipality);
        retVal.push(location);
        retVal.push(name);
        retVal.push(date);
        retVal.push(user);
        retVal.push(year);
        //console.log(retVal);
      }
    }
  });
  return retVal;
}

//readKFile("0206-840/3-(2019)");

function formFileName(row) {
  let retVal = "";
  let contractDate = row[3];
  let contractNumber = row[2].replaceAll("/", "_");
  let ownerName = row[4].replaceAll('"', "");
  let locationNumber = row[1];
  let municipality = row[0];
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
