const fs = require("fs");
const readLine = require("readline");
const Excel = require("exceljs");

const maximumRows = 1000000; // excels row limit per sheet

module.exports = (sheetName, sheet, fileName, workbook) => {
  let currentSheet = sheet;
  let lineCount = 0;
  let columns = [];
  let newSheetCount = 1;

  return new Promise((resolve, reject) => {
    const fileInStream = fs.createReadStream(fileName);
    const rl = readLine.createInterface({
      input: fileInStream,
      crlfDelay: Infinity,
    });
    rl.on("line", async (line) => {
      const toPersist = line.split(/\t/).map((el) => {
        if (!Number(el)) {
          return el;
        }
        return Number(el);
      });

      if (lineCount >= maximumRows) {
        console.log(`reached maximum row count, reseting to 0`);
        lineCount = 0;
        // commit the current sheet
        currentSheet.commit();
        // create a new sheet and set as the current
        currentSheet = workbook.addWorksheet(`${sheetName}-${++newSheetCount}`);

        // add the columns
        currentSheet.addRow(columns).commit();

        // continue as per usual now with this line data
        console.log(`reached maximum row count,created new sheet ${sheetName}-${newSheetCount}`);
        ++lineCount;
      } else {
        // console.log(`currentLineCount: ${lineCount}`);

        if (!columns.length) {
          console.log(
            `Adding columns none found -data: lineCount:${lineCount}, columns:${columns.length}`
          );
          // this is the column
          columns = toPersist.slice();
        }

        currentSheet.addRow(toPersist).commit();
        ++lineCount;
      }
    });

    rl.on("close", async () => {
      console.log(`closed file read`);
      await currentSheet.commit(); // close the worksheet
      resolve();
    });

    rl.on("error", (e) => {
      console.error(`Failed somewhere: `, e);
      fileInStream.destroy();
      reject();
    });
  });
};
