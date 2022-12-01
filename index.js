import readXlsxFile from "read-excel-file";

// File path.
readXlsxFile("/path/to/file").then((rows) => {
  rows.forEach((documentRow) => {
    console.log(documentRow);
  });

  // `rows` is an array of rows
  // each row being an array of cells.
});
