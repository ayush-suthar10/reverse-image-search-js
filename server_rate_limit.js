var fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');
const rateLimit = require('express-rate-limit');

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Results');

worksheet.columns = [
  { header: 'Image Name', key: 'imageName', width: 30 },
  { header: 'Result Link', key: 'link', width: 70 },
];

const resultsArray = [];

var express = require('express');
var app = express();
var http = require('http').Server(app);
var io = require('socket.io')(http);
var exec = require('child_process').exec;
var tmp = require('tmp');

tmp.setGracefulCleanup();

app.set('port', (process.env.PORT || 9001));
app.use(express.static('static'));

var uploadurl = 'http://images.google.com/searchbyimage/upload';

// Rate limiting middleware
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // max 100 requests per 15 minutes
  message: 'Too many requests from this IP, please try again after 15 minutes',
});
// this is a commit test
app.use(limiter);

let totalFiles = 0;
let processedFiles = 0;

io.on('connection', function (socket) {
  socket.on('image_upload_start', function(data) {
    totalFiles = data.totalFiles;
    processedFiles = 0;
    socket.emit('progress', { processed: processedFiles, total: totalFiles });
  });

  socket.on('image_upload', function (data) {
    var opts = { postfix: '.' + data['fext'] };
    tmp.file(opts, function (err, path, fd, cleanupCallback) {
      if (err) throw err;

      // Ensure buffer (convert if base64 string)
      var buffer = Buffer.isBuffer(data['fbuf']) ? data['fbuf'] : Buffer.from(data['fbuf'], 'base64');

      // Write buffer to the temp file path (NOT fd)
      fs.writeFile(path, buffer, function (err) {
        if (err) {
          cleanupCallback();
          throw err;
        }

        var cmd = ['curl', '-s', '-F', 'encoded_image=@' + path, uploadurl].join(' ');
        exec(cmd, function (err, stdout, stderr) {
          console.log('Curl output for', data['fname'], ':', stdout);
          var arr = /<A HREF="([^"]+)"/i.exec(stdout);
          if (arr != null) {
            resultsArray.push({ imageName: data['fname'], link: arr[1] });
            socket.emit('image_search', { url: arr[1], fname: data['fname'] });
          } else {
            console.log('No link found in curl output for:', data['fname']);
          }
          processedFiles++;
          socket.emit('progress', { processed: processedFiles, total: totalFiles });
          cleanupCallback();
        });
      });
    });
  });
});

app.get('/save-excel', async (req, res) => {
  worksheet.spliceRows(2, worksheet.rowCount - 1);
  resultsArray.forEach(item => {
    worksheet.addRow({
      imageName: item.imageName,
      link: item.link,
    });
  });
  await workbook.xlsx.writeFile('reverse_search_results.xlsx');
  console.log('Excel file saved');
  res.send('Excel file saved!');
});

http.listen(app.get('port'));
