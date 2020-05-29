'use strict';

var fs = require('fs'),
    express = require('express'),
    https = require('https');

var https_options = {
  key: fs.readFileSync(__dirname +'/../ssl-key.pem'),
  cert: fs.readFileSync(__dirname +'/../ssl-cert.pem')
};

var PORT = 3000,
    HOST = 'localhost';

var app = express();

// setup express to have static resource folders
app.use('/', express.static(__dirname + '/../public'));

var server = https.createServer(https_options, app)
                  .listen(PORT, HOST);

console.log('+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+');
console.log('HTTPS Server listening @ https://%s:%s', HOST, PORT);
console.log('+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+');
