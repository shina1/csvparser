// where the whole api starts
const express = require("express");

const bodyParser = require("body-parser");
const multer = require("multer");
let fs = require("fs");
const xlstojson = require("xls-to-json-lc");
const xlsxtojson = require("xlsx-to-json-lc");
const csvtojson = require("csvtojson");


const app = express();
app.use(bodyParser.json());
let storage = multer.diskStorage({
  //multers disk storage settings
  destination: function (req, file, cb) {
    cb(null, "./uploads/");
  },
  filename: function (req, file, cb) {
    const datetimestamp = Date.now();
    cb(
      null,
      file.fieldname +
        "-" +
        datetimestamp +
        "." +
        file.originalname.split(".")[file.originalname.split(".").length - 1]
    );
  },
});
// modify the upload variable
let upload = multer({
  //multer settings
  storage: storage,
  fileFilter: function (req, file, callback) {
    //file filter for recieving only .xls and .xlsx format only
    if (
      ["xls", "xlsx", "csv"].indexOf(
        file.originalname.split(".")[file.originalname.split(".").length - 1]
      ) === -1
    ) {
      return callback(new Error("Wrong extension type"));
    }
    callback(null, true);
  },
}).single("file");
/** API path that will upload the files */
app.post("/upload", function (req, res) {
  let exceltojson;
  upload(req, res, function (err) {
    if (err) {
      res.json({ error_code: 1, err_desc: err });
      return;
    }
    /** Multer gives us file info in req.file object */
    if (!req.file) {
      res.json({ error_code: 1, err_desc: "No file passed" });
      return;
    }
    /**The API Checks the extension of the incoming file, differantiate between .xls and .xlsx and 
           use the appropriate module
         */
    if (
      req.file.originalname.split(".")[
        req.file.originalname.split(".").length - 1
      ] === "xlsx"
    ) {
      exceltojson = xlsxtojson;
    } else if(
      req.file.originalname.split(".")[
        req.file.originalname.split(".").length - 1
      ] === "xls"
    ) {
      exceltojson = xlstojson;
    }else{
      exceltojson = csvtojson;
    }
    try {
      exceltojson(
        {
          input: req.file.path,
          output: null,
          lowerCaseHeaders: true,
        },

        function (err, result) {
          if (err) {
            return res.json({ error_code: 1, err_desc: err, data: null });
          }
         
        const resultObjs = result.map((el) => [el.firstname, el.lastname, el.age] );
          
          fs.writeFile("output.json", `${JSON.stringify(resultObjs)}`, "utf-8", (err) => {
            console.log("writting the file");
          });
          res.json({ resultObjs });
        }
      );
    } catch (e) {
      res.json({ error_code: 1, err_desc: "Corupted excel file" });
     
    }
  });
});
// now we collect the file from an input
app.get("/", function (req, res) {
  res.sendFile(__dirname + "/form.html");
});
// port server
let port = process.env.PORT;

app.listen(port, function () {
  console.log(`listening on port ${port}`);
});
