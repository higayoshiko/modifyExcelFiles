const express = require("express");
const app = express();
const path = require('node:path');
const ejs = require("ejs");
// const bodyParser = require("body-parser");
const port = 8888;

app.set("views", path.join(__dirname, "views"));
app.set("view engine", "ejs");

app.use(express.json());
// app.use(bodyParser.urlencoded({
//   extended: true
// }));
app.use(express.urlencoded({ extended: true }));

app.use(express.static("public"));
app.use("/css", express.static(__dirname + "public"));

const homeRoute = require("./routes/editExcel");
app.use("/", homeRoute);

app.listen(port, function() {
  console.log("server started");
});
