var express = require('express');
var ntlm = require('express-ntlm');
var path = require('path');
var app = express(); // using this we can use commands, function of express in this (server.js) file
var mongojs = require('mongojs');
var db = mongojs('providerapp'); // Means which mongodb database & collection we are going to use
var bodyParser = require('body-parser');
var multer = require('multer');
var excel = require('exceljs');
var fs = require("fs");
var xlsx_to_json = require("xlsx-to-json");
var xls_to_json = require("xls-to-json");
var GenerateSchema = require('generate-schema');
var json2xls = require('json2xls');
var http = require('http');
var conform = require('conform');
var mongoose = require('mongoose');


// To test whether server is running correctly
/* app.get("/", function(req, res){
 res.send("Hello world from server.js");
 }); */

app.use(express.static(__dirname + "/public")); // express.static means we are telling the server to look for static file i.e. html,css,js etc.

app.use(bodyParser.json()); // To parse the body that we received from input

app.get('/', function (req, res) {
    res.sendFile(__dirname + "/index.html");
});

var storage = multer.diskStorage({//multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './uploads/');
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length - 1]);
    }
});

app.use(multer({dest: './uploads/', storage: storage}).single('file'));
var workbook = new excel.Workbook();

// listens for the POST request from the controller
app.post('/contactlist', function (req, res) {

    //console.log("===========" + req.file.path);

    xls_to_json({
        input: req.file.path,
        output: "output.json"
    }, function (err, result) {
        if (err) {
            console.error(err);
        } else {
            console.log(result);
            var collectionName = req.body.title.toLowerCase().replace(/ /g, '_');
            db.collection(collectionName).insert(result, function (err, doc) {
                console.log(err);
                //res.json(doc);
            });
        }
    });

    req.body.dt_id = process.env['USERNAME'];
    req.body.user_domain = process.env['USERDOMAIN'];
    req.body.computer_name = process.env['COMPUTERNAME'];
    req.body.logon_server = process.env['LOGONSERVER'];
    db.collection('contactlist').insert(req.body, function (err, doc) {
        console.log(err);
        res.json(doc);
    });
});

//This tells the server to listen for the get request for created contactlist throughout
app.get('/contactlist', function (req, res) {
    db.collection('contactlist').find(function (err, docs) {
        //console.log(docs);
        res.json(docs);
    });
});

app.delete('/contactlist/:id', function (req, res) {
    var id = req.params.id; // to get the value of id from url
    console.log(id);
    db.collection('contactlist').remove({_id: mongojs.ObjectId(id)}, function (err, doc) {
        res.json(doc);
    });
});

app.get('/contactlist/:id', function (req, res) {
    var id = req.params.id;
    console.log(id);
    db.collection('contactlist').findOne({_id: mongojs.ObjectId(id)}, function (err, doc) {
        res.json(doc);
    });
});

app.put('/contactlist/:id', function (req, res) {
    var id = req.params.id;
    console.log(req.body.name);
    db.collection('contactlist').findAndModify({
        query: {_id: mongojs.ObjectId(id)},
        update: {$set: {
                title: req.body.title,
                description: req.body.description
            }}, new : true}, function (err, doc) {
        res.json(doc);
    });
});

app.get('/providerlist', function (req, res) {
    db.listCollections(function (err, collections) {
        res.json(collections);
    });
});

app.get('/getcollectiondata/:collection', function (req, res) {
    console.log(req.collection);
//    db.collection(req.collection).find(function (err, docs) {
//        res.json(docs);
//    });
});

var typeMappings =
        {
            "String": String,
            "Number": Number,
            "Boolean": Boolean,
            "ObjectId": mongoose.Schema.ObjectId
        };

function makeSchema(jsonSchema) {
    var outputSchemaDef = {};
    for (fieldName in jsonSchema.data) {
        var fieldType = jsonSchema.data[fieldName];
        if (typeMappings[fieldType]) {
            outputSchemaDef[fieldName] = typeMappings[fieldType];
        } else {
            console.error("invalid type specified:", fieldType);
        }
    }
    return new mongoose.Schema(outputSchemaDef);
}

app.get('/providerlist/:id', function (req, res) {
    var id = req.params.id;
    db.collection('provider').findOne({_id: mongojs.ObjectId(id)}, function (err, data) {
        res.json(data);
    });
});

app.use(json2xls.middleware);

app.put('/export/:id', function (req, res) {
    var id = req.params.id;
    console.log(req.body.provider_name);
    db.collection('provider').findAndModify({
        query: {_id: mongojs.ObjectId(id)},
        update: {$set: {
                email: req.body.email,
                dt_number: req.body.dt_number,
                provider_name: req.body.provider_name,
                provider_dob: req.body.provider_dob,
                contact_number: req.body.contact_number,
                country: req.body.country,
                Department: req.body.Department
            }}, new : true}, function (err, doc) {

        var jsonArr = {};
        var jsonArr = doc;
        var xls = json2xls(jsonArr);
        fs.writeFileSync('data.xlsx', xls, 'binary');

        res.setHeader('Content-disposition', 'attachment; filename=data.xlsx');
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        res.xls('data.xlsx', jsonArr);

    });
});

/****************************************************************************************************************************************/




/****************************************************************************************************************************************/

app.listen(1000);
console.log("Server running on port 1000");
