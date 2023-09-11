const express = require("express");
const router = express.Router();
const https = require('node:https');
const xlsx = require('xlsx');


const workbook = xlsx.readFile("../sample.xlsx");
const sheet = workbook.Sheets["Sheet1"];

const range = xlsx.utils.decode_range(sheet["!ref"]);

try {

  for (let rowNumber = range.s.r + 1; rowNumber <= range.e.r; rowNumber++){

    if (sheet[xlsx.utils.encode_cell({r: rowNumber, c: 2})] !== undefined){
      let email = sheet[xlsx.utils.encode_cell({r: rowNumber, c: 2})].v;
          email = email.replace(/\s/g, '').toLowerCase().split("@");

      const commonEmailsArray = ["gmail.com", "yahoo.com", "i.softbank.jp", "icloud.com"]
          for(let commonEmail = 0; commonEmail <= commonEmailsArray.length; commonEmail++){

            if ((email[1] + commonEmailsArray[commonEmail]).split('').sort().join('').replace(/(.)\1+/g, "").length <= 1){
              sheet[xlsx.utils.encode_cell({r: rowNumber, c: 2})].v = email[0] + "@" + commonEmailsArray[commonEmail];

                  }
          }

    }

    const newFile = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newFile, sheet, "sheet1");
    xlsx.writeFile(newFile, "editedFile.xlsx");
  }

  }catch(err){
    console.log(err);
  }

module.exports = router;
