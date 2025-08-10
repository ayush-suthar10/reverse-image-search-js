var fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');

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

io.on('connection', function (socket) {
    socket.on('image_upload', function (data) {
        var opts = { postfix: '.' + data['fext'] };
        tmp.file(opts, function (err, path, fd, cleanupCallback) {
            if (err) throw err;
            var buffer = data['fbuf'];
            fs.writeSync(fd, buffer, 0, buffer.length);
            fs.close(fd, function (err) {
                if (err) throw err;
                var cmd = ['curl', '-s', '-F', 'encoded_image=@' + path, uploadurl].join(' ');
                exec(cmd, function (err, stdout, stderr) {
                    console.log('Curl output for', data['fname'], ':', stdout);  // DEBUG LOG
                    var arr = /<A HREF="([^"]+)"/i.exec(stdout); // Improved regex for redirect link
                    if (arr != null) {
                        resultsArray.push({ imageName: data['fname'], link: arr[1] });
                        socket.emit('image_search', { url: arr[1], fname: data['fname'] });
                    } else {
                        console.log('No link found in curl output for:', data['fname']);
                    }
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
