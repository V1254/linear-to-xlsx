const glob = require("glob");
const addToSheet = require("./processFile");
const Excel = require("exceljs");

const APO_TYPES = {
  // ADDITIVE: "_Additive.",
  // DOMINANT: "_Dominant.",
  GENOTYPIC: "_Genotypic.",
  // RECESIVE: "_Recessive.",
};
const BASE_PATH = "D:/v3/";
const OUTPUT_PATH = "D:/v3/parsed/";

// navigate to the path
process.chdir(BASE_PATH);

const wrapper = async () => {
  for await ([, val] of Object.entries(APO_TYPES)) {
    console.log(`Processing ${val} sheets`);
    const workbook = new Excel.stream.xlsx.WorkbookWriter({
      filename: OUTPUT_PATH + val.split("_")[1] + ".xlsx",
    });
    // const outputFileName = val.split("_")[1]; // done
    let index = 1;
    await glob(`*${val}*.linear`, async (err, files) => {
      for (const fileName of files) {
        console.log(`Processing ${index} of ${files.length}`);
        index++;
        const splitByDot = fileName.split("."); // Additive.transform.., .glim
        const sheetName = splitByDot[1]
          .replace(/transformed_EDU_/, "")
          .replace(/transformed_APO_/, "")
          .replace(/transformed_MIG_/, "");
        console.log(`Processing sheet : ${sheetName} of ${fileName}`);

        const sheet = workbook.addWorksheet(sheetName);
        await addToSheet(sheetName, sheet, fileName, workbook);

        // await workbook.commit();
      }

      //   const first2 = files.slice(0, 2);
      //   console.log(`files: `, files);
      //   await processFile(first2[1], OUTPUT_PATH + val + ".xlsx", "test");
      //   // await processFile(first2[2], OUTPUT_PATH + val + ".xlsx");
      //   console.log(`finished processing ??`);
      console.log(`Finished Processing all ${val} sheets, committing`);
      await workbook.commit();
    });
  }
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
