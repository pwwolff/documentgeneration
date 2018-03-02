// server.js

// BASE SETUP
// =============================================================================

// call the packages we need
var express = require('express');        // call express
var app = express();                 // define our app using express
var bodyParser = require('body-parser');
var XlsxTemplate = require('xlsx-template')
var fs = require('fs')
var path = require('path')
var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
// configure app to use bodyParser()
// this will let us get the data from a POST
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

var port = process.env.PORT || 8080;        // set our port

// ROUTES FOR OUR API
// =============================================================================
var router = express.Router();              // get an instance of the express Router

// test route to make sure everything is working (accessed at GET http://localhost:8080/api)
router.post('/xlsx/:templateName', function (req, res) {
    const templateName = req.params.templateName
    const values = req.body

    // Load an XLSX file into memory
    fs.readFile(path.join(__dirname, 'templates', templateName + '.xlsx'), function (err, data) {

        // Create a template
        var template = new XlsxTemplate(data);

        // Replacements take place on first sheet
        var sheetNumber = 1;

        // Perform substitution
        template.substitute(sheetNumber, values);

        // Get binary data
        var data = template.generate();

        fs.writeFileSync(path.join(__dirname, 'templates', 'output.xlsx'), data, 'binary', function (err) {
            if (err) {
                res.json({ message: 'Error: ' + err });
            }
            else {
                res.json({ message: 'Success!'});
            }
        });
    });
});

// test route to make sure everything is working (accessed at GET http://localhost:8080/api)
router.post('/docx/:templateName', function (req, res) {
    const templateName = req.params.templateName
    const values = req.body
    
    const content = fs.readFileSync(path.join(__dirname, 'templates',  templateName + '.docx'), 'binary');

    var zip = new JSZip(content);

    var doc = new Docxtemplater();
    doc.loadZip(zip);
    
    //set the templateVariables
    doc.setData({values});

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render()
    }
    catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }
        console.log(JSON.stringify({error: e}));
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
        throw error;
    }
    
    var buf = doc.getZip()
                 .generate({type: 'nodebuffer'});
    
    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.join(__dirname, 'templates', 'output.docx'), buf);
    res.json({message: "Success"})
});

// REGISTER OUR ROUTES -------------------------------
// all of our routes will be prefixed with /api
app.use('/api', router);

// START THE SERVER
// =============================================================================
app.listen(port);
console.log('Magic happens on port ' + port);
