var createError = require("http-errors");
var express = require("express");
var path = require("path");
var cookieParser = require("cookie-parser");
const cors = require("cors");
const bodyParser = require("body-parser");
var morgan = require("morgan");
var fs = require("fs");
const rfs = require("rotating-file-stream");
const basicAuth = require("express-basic-auth");
const _config = require("./appSetting.js");
const databaseConnect = require("./dbconnect.js");
const moment = require('moment');



var indexRouter = require("./routes/index");
var reportRoute = require("./routes/report.js");
const gendocRoute = require("./routes/gendoc.js");
var app = express();


/**
 * 
 * Log Control
 */

var fileUpload = require("express-fileupload");
const { log } = require("console");
app.use(
  fileUpload({
    createParentPath: true,
  })
);






app.set("views", path.join(__dirname, "views"));
app.set("view engine", "pug");
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb' }));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, "public")));
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));



app.use("/", indexRouter);

app.use("/api/report", reportRoute);
app.use("/api/gendoc", gendocRoute);

console.log('Report Service is running on port 4301');


app.use(function (req, res, next) {
  next(createError(404));
});

app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get("env") === "development" ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render("error");
});

app.disable("x-powered-by");

module.exports = app;
