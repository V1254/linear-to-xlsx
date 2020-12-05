const glob = require("glob");
const addToSheet = require("./processFile");
const Excel = require("exceljs");

const types = ["_Additive.", "_Dominant.", "_Genotypic.", "_Recessive."];
// const types = ["_Genotypic."];
const BASE_PATH = "D:/v3/";
const OUTPUT_PATH = "D:/v3/parsed/";

// navigate to the path
process.chdir(BASE_PATH);

const globWrapper = (filesToMatch, currentFileName) =>
  new Promise((resolve, resject) => {
    glob(`*${filesToMatch}*.linear`, async (err, files) => {
      let currentWorkBook = new Excel.stream.xlsx.WorkbookWriter({
        filename: OUTPUT_PATH + currentFileName + ".xlsx",
      });
      let index = 1;
      const _gentypicCount = 4; // only fit 3 genotypic files per book so it doesn't get huge
      let _createGntypicCount = 0;
      for (const fileName of files) {
        console.log(`\nProcessing ${index} of ${files.length}\n`);
        index++;
        const splitByDot = fileName.split("."); // Additive.transform.., .glim
        const sheetName = splitByDot[1]
          .replace(/transformed_EDU_/, "")
          .replace(/transformed_APO_/, "")
          .replace(/transformed_MIG_/, "");
        if (filesToMatch === "_Genotypic." && _createGntypicCount >= _gentypicCount) {
          console.log(
            `Hit max Files for current workbook [${_createGntypicCount}|${_gentypicCount}] saving current workbook: ${currentFileName}.xlsx`
          );
          await currentWorkBook.commit();

          console.log(`Finished saving ${currentFileName}.xlsx`);

          // nextFileName
          const splitCurrentFileName = currentFileName.split("-");
          if (splitCurrentFileName.length === 1) {
            // first addition we are doing
            currentFileName = currentFileName + "-1";
          } else {
            currentFileName =
              splitCurrentFileName[0] + "-" + (parseInt(splitCurrentFileName[1]) + 1);
          }

          // new workbook
          currentWorkBook = new Excel.stream.xlsx.WorkbookWriter({
            filename: OUTPUT_PATH + currentFileName + ".xlsx",
          });

          _createGntypicCount = 0;
        }

        console.log(`Adding file: ${fileName} to ${currentFileName}.xlsx`);

        const sheet = currentWorkBook.addWorksheet(sheetName);
        await addToSheet(sheetName, sheet, fileName, currentWorkBook);

        console.log(`\n\n======================================\n`);

        // increment the count after adding to the workbook for this file
        if (filesToMatch === "_Genotypic.") {
          _createGntypicCount++;
        }
      }
      console.log(`Finished Processing all ${filesToMatch} sheets, committing`);
      await currentWorkBook.commit();
      resolve();
    });
  });

const wrapper = async () => {
  for (let i = 0; i < types.length; i++) {
    const currentType = types[i];
    let currentFileName = currentType.split("_")[1].replace(".", "");
    console.log(`======================================`);
    console.log(`\nStarting Work on ${currentType} files \n`);
    // let currentWorkBook = new Excel.stream.xlsx.WorkbookWriter({
    //   filename: OUTPUT_PATH + currentFileName + ".xlsx",
    // });

    await globWrapper(currentType, currentFileName);
  }

  // for await ([, val] of Object.entries(APO_TYPES)) {
  //   let currentFileName = val.split("_")[1].replace(".", "");
  //   console.log(`Processing ${val} sheets`);
  //   let currentWorkBook = new Excel.stream.xlsx.WorkbookWriter({
  //     filename: OUTPUT_PATH + currentFileName + ".xlsx",
  //   });
  //   // const outputFileName = val.split("_")[1]; // done
  //   let index = 1;
  //   const _gentypicCount = 3; // only fit 3 genotypic files per book so it doesn't get huge
  //   let _createGntypicCount = 0;
  // }
};

wrapper();

// const fs = require("fs");
// const readLine = require("readline");
// const Excel = require("exceljs");

// const path = `D:/v3/chr4_APO_Additive.transformed_APO_FRAG_30.glm.linear`;

// // sheet name will be APO_FRAG_30
// // Additive

// const workbook = new Excel.stream.xlsx.WorkbookWriter({
//   filename: "D:/v3/parsed/first.xlsx",
// });

// const sheet = workbook.addWorksheet("My Sheet");

// const fileInStream = fs.createReadStream(path);
// const rl = readLine.createInterface({
//   input: fileInStream,
//   crlfDelay: Infinity,
// });

// const read = async () => {
//   for await (const chunk of fileInStream) {
//     console.log(chunk.toString());
//     const data = chunk.toString().split(/\t/).map((el) => {
//       if (!Number(el)) {
//         return el;
//       }

//       return Number(el);
//     });

//     sheet.addRow(data).commit();
//   }

//   await workbook.commit();
// };

// read();

// // let columned = false;
// // let id = 1;
// // rl.on("line", (line) => {
// //   const toPersist = line.split(/\t/).map((el) => {
// //     if (!Number(el)) {
// //       return el;
// //     }

// //     return Number(el);
// //   });
// //   sheet.addRow(toPersist).commit();
// // });

// // rl.on("close", async () => {
// //   console.log(`finished reading file, commting the excel`);
// //   await workbook.commit();
// // });

// // fs.readFile(path, "utf-8", (err, data) => {
// //   if (err) return console.log(`failed: `, err);
// //   console.log(data);
// // });
