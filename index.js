require("dotenv").config();
const glob = require("glob");
const addToSheet = require("./writeToSheet");
const Excel = require("exceljs");

const types = ["_Additive.", "_Dominant.", "_Genotypic.", "_Recessive."];
const BASE_PATH = process.env.PATH_TO_FILES;
const OUTPUT_PATH = process.env.PATH_TO_OUTPUT;

// navigate to the path
process.chdir(BASE_PATH);

const generateBooks = (filesToMatch, currentFileName) =>
  new Promise((resolve, resject) => {
    glob(`*${filesToMatch}*.linear`, async (err, files) => {
      let currentWorkBook = new Excel.stream.xlsx.WorkbookWriter({
        filename: OUTPUT_PATH + currentFileName + ".xlsx",
      });
      let index = 1;
      const _gentypicCount = parseInt(process.env.MAX_FILE_PER_BOOK_Genotypic); // only fit 3 genotypic files per book so it doesn't get huge
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

const main = async () => {
  for (let i = 0; i < types.length; i++) {
    const currentType = types[i];
    let currentFileName = currentType.split("_")[1].replace(".", "");
    console.log(`======================================`);
    console.log(`\nStarting Work on ${currentType} files \n`);
    // let currentWorkBook = new Excel.stream.xlsx.WorkbookWriter({
    //   filename: OUTPUT_PATH + currentFileName + ".xlsx",
    // });

    await generateBooks(currentType, currentFileName);
  }
};

main();
