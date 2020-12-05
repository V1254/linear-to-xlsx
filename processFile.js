const fs = require("fs");
const readLine = require("readline");
const Excel = require("exceljs");

const maximumRows = 1048576; // excels row limit per sheet

module.exports = (sheetName, sheet, fileName, workbook) =>
  new Promise(async (resolve, reject) => {
    const fileInStream = fs.createReadStream(fileName);
    const rl = readLine.createInterface({
      input: fileInStream,
      crlfDelay: Infinity,
    });

    let currentSheet = sheet;
    let lineCount = 0;
    let columns = [];
    let newSheetCount = 1;
    rl.on("line", (line) => {
      if (lineCount === maximumRows) {
        // commit the current sheet
        currentSheet.commit();
        // create a new sheet and set as the current
        currentSheet = workbook.addWorksheet(`${sheetName}-${++newSheetCount}`);

        // add the columns
        currentSheet.addRow(columns).commit();

        // reset the counters
        newSheetCount++;
        lineCount++;
        // continue as per usual now with this line data
      }

      const toPersist = line.split(/\t/).map((el) => {
        if (!Number(el)) {
          return el;
        }
        return Number(el);
      });

      if (lineCount === 0) {
        // this is the column
        columns = toPersist.slice();
      }

      currentSheet.addRow(toPersist).commit();
      lineCount++;
    });

    rl.on("close", async () => {
      console.log(`closed file read`);
      fileInStream.destroy();
      currentSheet.commit(); // close the worksheet
      resolve();
    });

    rl.on("error", (e) => {
      console.error(`Failed somewhere: `, e);
      fileInStream.destroy();
      reject();
    });
  });
