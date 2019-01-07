var fs = require('fs');
var archiver = require('archiver');

var archive = archiver('zip', {
    zlib: { level: 9 } // Sets the compression level.
});

// append files from a sub-directory, putting its contents at the root of archive
archive.directory(__dirname + '/../dist/', false);

// pipe archive data to the file
var output = fs.createWriteStream(__dirname + '/../dist/dist.zip');
archive.pipe(output);
archive.finalize();
