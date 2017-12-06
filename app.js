// Run with app.exe --port=3000 --secret=ed1bb543-8767-4ab3-ad94-ad6db1c225b5
var port = 3000;
var secret = "";
var debugLogEnabled = false;

process.argv.forEach((val, index) => {
    if (val.startsWith('--port='))
        port = parseInt(val.split('=')[1]);
    else if (val.startsWith('--secret='))
        secret = val.split('=')[1];
    else if (val === '--verbose')
        debugLogEnabled = true;
});

console.log('Running Exchange Web Service Proxy on http://localhost:' + port + ' using secret ' + secret);
if (debugLogEnabled) {
    console.log('  => Verbose Logging active');
}

var app = require('./express-app');
app.configureApp(secret, debugLogEnabled);
app.listen(port);