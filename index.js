const glob = require("glob");
const fs = require("fs");
const readLine = require("readline");
const officegen = require("officegen");

const APO_TYPES = {
  ADDITIVE: "APO_Additive",
  //   DOMINANT: "APO_Dominant",
  //   GENOTYPIC: "APO_Genotypic",
  //   RECESIVE: "APO_Recessive",
};
const BASE_PATH = "D:/v3/";
const OUTPUT_PATH = "D:/v3/parsed/";

process.chdir(BASE_PATH); // navigate to the path above so we can pattern match
Object.keys(APO_TYPES).forEach((type) => {
  const stringToMatch = APO_TYPES[type];
  glob(`*${stringToMatch}*.linear`, (err, files) => {
    console.info(`Working on ${stringToMatch} Sheets`);
    files.forEach(async (file, index) => {
      if (index !== 0) return;

      const fileInStream = fs.createReadStream(file);
      const rl = readLine.createInterface({
        input: fileInStream,
        crlfDelay: Infinity,
      });

      const fileOutStream = fs.createWriteStream(`${OUTPUT_PATH}${stringToMatch}.xlsx`);

      const xlsx = officegen("xlsx");

      //   xlsx.on("finalize", (writter) => console.log(`Finished writing to spreadhseet`));

      xlsx.on("error", (err) => console.log(`failed somewhere `, err));

      const sheet = xlsx.makeNewSheet();
      sheet.name = stringToMatch;

      let columns = [];
      let currentData = [];
      let columnIndex = 0;
      let rowCount = 0;
      rl.on("line", (lineData) => {
        if (!columnIndex) {
          // first column and hence the columns
          console.log(`Columns is: ${lineData}`);
          columns = lineData.split(/\t/);
          // A1 (65) Z1
          //   let startCharCode = 65;
          //   const startRow = 1;
          //   const excelColumns = new Array(columns.length).map((el, index) => {

          //   })

          // create the columns in the spreadsheet
          let startCharCode = 65; // letter A = 65 , B=66 etc...
          columns.forEach((column) => {
            const toChar = String.fromCharCode(startCharCode);
            sheet.setCell(`${toChar}1`, column);
            startCharCode++;
          });
          xlsx.generate(fileOutStream);
          columnIndex++;
        }

        if (rowCount > 10) return;
        //
        rowCount++;
        currentData = lineData.split(/\t/);
        let startCharCode = 65; // letter A = 65 , B=66 etc...
        currentData.forEach((data) => {
          const toChar = String.fromCharCode(startCharCode);
          sheet.setCell(`${toChar}1`, data);
          startCharCode++;
        });
        xlsx.generate(fileOutStream);
      });

      //   fs.readFile(value, (err, data) => {
      //     if (!err && data) {
      //       console.log(`data: `, data);
      //     }
      //   });
    });

    rl;
  });
});

// glob("D:/v3/*.linear", (err, files) => console.log(`errro: ${err}`, files));
