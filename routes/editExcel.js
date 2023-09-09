const express = require("express");
const router = express.Router();
const https = require('node:https');
const xlsx = require('xlsx');


// router.post("/", function(req, res) {
//   // let excelFile = req.body.excelFile;
//   // let sheetName = req.body.excelFile;
//
//   const workbook = xlsx.readFile("sample.xlsx");
//   const sheet = workbook.Sheets("Sheet1");
//
//   const range = xlsx.utils.decode_range(worksheet["!ref"]);
//
//   try {
//
//     for (let rowNumber = range.s.r;rowNumber <= range.e.r; rowNumber++){
//       const email = sheet[xlsx.utils.encode_cell({r: rowNumber, c: 3})].v; //grabbing the email cell
//
//       if (email === "yushiko444@yahoo.com"){
//         sheet[xlsx.utils.encode_cell({r: rowNumber, c: 3})].v + "jello";
//       }
//
//       const newFile = xlsx.utils.book_new();
//       xlsx.utils.book_append_sheet(newFile, sheet, "sheet1");
//       xlsx.writeFile(newFile, "editedFile.xlsx");
//     }
//
//     }catch(err){
//       console.log(err);
//     }
// });


const workbook = xlsx.readFile("../sample.xlsx");
const sheet = workbook.Sheets["Sheet1"];

const range = xlsx.utils.decode_range(sheet["!ref"]);

try {

  for (let rowNumber = range.s.r; rowNumber <= range.e.r; rowNumber++){
    const email = sheet[xlsx.utils.encode_cell({r: rowNumber, c: 3})];
    //grabbing the email cell

    // console.log(email)
    console.log(range.s.r, rowNumber)
    console.log(range.e.r)

    // if (email === "yushiko444@yahoo.com"){
    //   sheet[xlsx.utils.encode_cell({r: rowNumber, c: 3})].v + "jello";
    // }

    const newFile = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newFile, sheet, "sheet1");
    xlsx.writeFile(newFile, "editedFile.xlsx");
  }

  }catch(err){
    console.log(err);
  }

module.exports = router;
