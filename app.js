// Generated by CoffeeScript 1.7.1
(function() {
  var ADDR1_COLUMN, ADDR2_COLUMN, ADDR3_COLUMN, BANK1_COLUMN, BANK2_COLUMN, BANK3_COLUMN, BANK4_COLUMN, COLUMNS_WIDTH, ID_COLUMN, NUM_COLUMN, OUTPUT_PATH, PRICE_COLUMN, SpreadsheetReader, TAX_COLUMN, WRITER_COLUMN, app, busboy, express, formidable, fs, makeBill, officegen, path, readFile, util, xlsx;

  express = require("express");

  busboy = require("connect-busboy");

  formidable = require("formidable");

  util = require("util");

  path = require("path");

  app = express();

  app.set('view engine', 'jade');

  app.set('views', __dirname + '/views');

  app.use(busboy());

  fs = require("fs");

  officegen = require("officegen");

  xlsx = void 0;

  SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

  OUTPUT_PATH = "writer-bill.xlsx";

  WRITER_COLUMN = 3;

  TAX_COLUMN = 4;

  PRICE_COLUMN = 5;

  NUM_COLUMN = 6;

  ID_COLUMN = 7;

  BANK1_COLUMN = 8;

  BANK2_COLUMN = 9;

  BANK3_COLUMN = 10;

  BANK4_COLUMN = 11;

  ADDR1_COLUMN = 12;

  ADDR2_COLUMN = 13;

  ADDR3_COLUMN = 14;

  COLUMNS_WIDTH = 2.2;

  app.get('/', function(req, res) {
    return res.render('index');
  });

  app.post('/upload', function(req, res) {
    var form;
    xlsx = officegen('xlsx');
    form = new formidable.IncomingForm();
    return form.parse(req, function(err, fields, files) {
      var file_ext, file_name, file_size, index, new_path, old_path;
      old_path = files.file.path;
      file_size = files.file.size;
      file_ext = files.file.name.split('.').pop();
      index = old_path.lastIndexOf('/') + 1;
      file_name = old_path.substr(index);
      new_path = path.join(process.env.PWD, '/files/', file_name + '.' + file_ext);
      return fs.readFile(old_path, function(err, data) {
        return fs.writeFile(new_path, data, function(err) {
          return fs.unlink(old_path, function(err) {
            if (err) {
              res.status(500);
              return res.json({
                'success': false
              });
            } else {
              res.status(200);
              readFile(new_path);
              return fs.readFile(OUTPUT_PATH, function(err, data) {
                res.writeHead(200, {
                  'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });
                return res.end(data, 'binary');
              });
            }
          });
        });
      });
    });
  });

  app.listen(3000);

  readFile = function(filePath) {
    return SpreadsheetReader.read(filePath, function(err, workbook) {
      var out, writers;
      if (err) {
        console.log(err);
        return;
      }
      writers = [];
      workbook.sheets.forEach(function(sheet) {
        var stopLoop;
        console.log('sheet: %s(%d)', sheet.name, sheet.index);
        if (sheet.index !== 0) {
          return;
        }
        console.log('sheet: %s(%d)', sheet.name, sheet.index);
        stopLoop = false;
        return sheet.rows.forEach(function(row) {
          var writer;
          if (stopLoop) {
            return;
          }
          writer = {};
          row.forEach(function(cell) {
            if (cell.row >= 4 && !stopLoop) {
              switch (cell.column) {
                case WRITER_COLUMN:
                  writer["name"] = cell.value;
                  break;
                case TAX_COLUMN:
                  if (cell.value === "個人") {
                    writer["tax_name"] = "源泉徴収税";
                    writer["tax_val"] = 10.21;
                  } else if (cell.value === "法人") {
                    writer["tax_name"] = "消費税";
                    writer["tax_val"] = 8.00;
                  }
                  break;
                case PRICE_COLUMN:
                  writer["price"] = cell.value;
                  break;
                case NUM_COLUMN:
                  writer["num"] = cell.value;
                  if (cell.value && cell.value > 0) {
                    writer["sum"] = writer["price"] * writer["num"];
                  }
                  break;
                case ID_COLUMN:
                  if (cell.value) {
                    writer["id"] = cell.value;
                  }
                  break;
                case BANK1_COLUMN:
                  if (cell.value) {
                    writer["bank1"] = cell.value;
                  }
                  break;
                case BANK2_COLUMN:
                  if (cell.value) {
                    writer["bank2"] = cell.value;
                  }
                  break;
                case BANK3_COLUMN:
                  if (cell.value) {
                    writer["bank3"] = cell.value;
                  }
                  break;
                case BANK4_COLUMN:
                  if (cell.value) {
                    writer["bank4"] = cell.value;
                  }
                  break;
                case ADDR1_COLUMN:
                  if (cell.value) {
                    writer["addr1"] = cell.value;
                  }
                  break;
                case ADDR2_COLUMN:
                  if (cell.value) {
                    writer["addr2"] = cell.value;
                  }
                  break;
                case ADDR3_COLUMN:
                  if (cell.value) {
                    writer["addr3"] = cell.value;
                  }
              }
              if (cell.column > ADDR3_COLUMN) {
                if (!writer["name"] || writer["name"] === "合計") {
                  stopLoop = true;
                }
              }
            }
          });
          if (writer["name"] && writer["num"] && writer["num"] > 0) {
            writers.push(writer);
            return console.log(writer);
          }
        });
      });
      writers.forEach(function(writer) {
        if (writer["name"] && writer["num"] && writer["num"] > 0 && writer["sum"] && writer["sum"] > 0) {
          return makeBill(writer);
        }
      });
      out = fs.createWriteStream(OUTPUT_PATH);
      out.on('error', function(err) {
        return console.log("Error: " + err);
      });
      return xlsx.generate(out, {
        finalize: function(written) {
          return console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n');
        },
        error: function(err) {
          return console.log(err);
        }
      });
    });
  };

  makeBill = function(writer) {
    var dd, i, j, mon, sheet, yy, _i, _j, _k, _l;
    sheet = xlsx.makeNewSheet();
    sheet.name = writer["name"];
    sheet.columnsWidth = [];
    for (i = _i = 0; _i <= 21; i = ++_i) {
      sheet.columnsWidth[i] = COLUMNS_WIDTH;
    }
    sheet.setCellWithStyle('B2', '御　請　求　書', '20B');
    sheet.mergeCells([3, 1], [3, 6]);
    sheet.setCellWithStyle('B4', '株式会社 Ｄｏｎｕｔｓ 御中', '18BU');
    sheet.setCellWithStyle('C4', 'a', '18BU');
    sheet.setCellWithStyle('D4', 'a', '18BU');
    sheet.setCellWithStyle('E4', 'a', '18BU');
    sheet.setCellWithStyle('F4', 'a', '18BU');
    sheet.setCellWithStyle('G4', 'a', '18BU');
    dd = new Date();
    yy = dd.getYear();
    if (yy < 2000) {
      yy += 1900;
    }
    yy = yy - 2000 + 12;
    mon = dd.getMonth() + 1;
    dd = dd.getDate();
    sheet.setCellWithStyle('O2', '平成' + yy + '年' + mon + '月' + dd + '日', '14BU');
    sheet.setCellWithStyle('R2', 'a', '14BU');
    sheet.setCellWithStyle('S2', 'a', '14BU');
    sheet.setCellWithStyle('T2', 'a', '14BU');
    if (writer["id"]) {
      sheet.setCellWithStyle('O1', '請求番号: ' + writer["id"], '11');
    }
    sheet.setCellWithStyle('B6', '下記のとおり御請求申し上げます', '12');
    if (writer["addr1"]) {
      sheet.setCellWithStyle('O7', writer["addr1"], '12');
    }
    if (writer["addr2"]) {
      sheet.setCellWithStyle('O8', writer["addr2"], '12');
    }
    if (writer["addr3"]) {
      sheet.setCellWithStyle('O9', writer["addr3"], '12');
    }
    sheet.setCellWithStyle('O10', writer["name"], '12');
    sheet.setCellWithStyle('B11', '振込先銀行', '14BU');
    sheet.setCellWithStyle('C11', 'a', '14BU');
    sheet.setCellWithStyle('D11', 'a', '14BU');
    sheet.setCellWithStyle('B13', '口座番号', '14BU');
    sheet.setCellWithStyle('C13', 'a', '14BU');
    sheet.setCellWithStyle('D13', 'a', '14BU');
    sheet.setCellWithStyle('B14', '名義', '14BU');
    sheet.setCellWithStyle('C14', 'a', '14BU');
    sheet.setCellWithStyle('D14', 'a', '14BU');
    if (writer["bank1"]) {
      sheet.setCellWithStyle('E11', writer["bank1"], '14BU');
    }
    sheet.setCellWithStyle('F11', 'a', '14BU');
    sheet.setCellWithStyle('G11', 'a', '14BU');
    sheet.setCellWithStyle('H11', 'a', '14BU');
    sheet.setCellWithStyle('I11', 'a', '14BU');
    if (writer["bank2"]) {
      sheet.setCellWithStyle('E12', writer["bank2"], '14BU');
    }
    sheet.setCellWithStyle('F12', 'a', '14BU');
    sheet.setCellWithStyle('G12', 'a', '14BU');
    sheet.setCellWithStyle('H12', 'a', '14BU');
    sheet.setCellWithStyle('I12', 'a', '14BU');
    if (writer["bank3"]) {
      sheet.setCellWithStyle('E13', writer["bank3"], '14BU');
    }
    sheet.setCellWithStyle('F13', 'a', '14BU');
    sheet.setCellWithStyle('G13', 'a', '14BU');
    sheet.setCellWithStyle('H13', 'a', '14BU');
    sheet.setCellWithStyle('I13', 'a', '14BU');
    if (writer["bank4"]) {
      sheet.setCellWithStyle('E14', writer["bank4"], '14BU');
    }
    sheet.setCellWithStyle('F14', 'a', '14BU');
    sheet.setCellWithStyle('G14', 'a', '14BU');
    sheet.setCellWithStyle('H14', 'a', '14BU');
    sheet.setCellWithStyle('I14', 'a', '14BU');
    sheet.setCellWithStyle('B16', '合計金額', '24C');
    sheet.setCellWithStyle('C16', 'a', '24C');
    sheet.setCellWithStyle('D16', 'a', '24C');
    sheet.setCellWithStyle('E16', 'a', '24C');
    sheet.setCellWithStyle('F16', 'a', '24C');
    sheet.setCellWithStyle('G16', 'a', '24C');
    sheet.setCellWithStyle('H16', 'a', '24C');
    sheet.setCellWithStyle('I16', '¥' + String(writer["sum"] - Math.ceil(writer["sum"] * writer["tax_val"] / 100)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,'), '22C');
    sheet.setCellWithStyle('J16', 'a', '22C');
    sheet.setCellWithStyle('K16', 'a', '22C');
    sheet.setCellWithStyle('L16', 'a', '22C');
    sheet.setCellWithStyle('M16', ' ', 'mdB');
    sheet.setCellWithStyle('N16', ' ', 'mdB');
    sheet.setCellWithStyle('O16', ' ', 'mdB');
    sheet.setCellWithStyle('P16', ' ', 'mdB');
    sheet.setCellWithStyle('Q16', ' ', 'mdB');
    sheet.setCellWithStyle('R16', ' ', 'mdB');
    sheet.setCellWithStyle('S16', ' ', 'mdB');
    sheet.setCellWithStyle('T16', ' ', 'mdB');
    sheet.setCellWithStyle('U16', ' ', 'mdB');
    for (i = _j = 17; _j <= 30; i = ++_j) {
      for (j = _k = 1; _k <= 20; j = ++_k) {
        sheet.setCellWithStyle(String.fromCharCode('A'.charCodeAt(0) + j) + i, ' ', 'thB');
        if (j === 1 || j === 7 || j === 9 || j === 12 || j === 17) {
          if (j !== 1 || i !== 30) {
            if (j === 1 && i - 17 <= 10) {
              sheet.setCellWithStyle(String.fromCharCode('A'.charCodeAt(0) + j) + i, '' + (i - 17), 'thBR');
            } else {
              sheet.setCellWithStyle(String.fromCharCode('A'.charCodeAt(0) + j) + i, ' ', 'thBR');
            }
          }
        }
      }
    }
    sheet.setCellWithStyle('B17', '摘要', '12C');
    sheet.setCellWithStyle('I17', '数量', '12C');
    sheet.setCellWithStyle('K17', '単価', '12C');
    sheet.setCellWithStyle('N17', '金額(税抜)', '12C');
    sheet.setCellWithStyle('S17', '備考', '12C');
    sheet.setCellWithStyle('C18', '原稿料(' + mon + '月分)', '12C');
    sheet.setCellWithStyle('I18', '' + writer["num"], '12C');
    sheet.setCellWithStyle('K18', '¥' + String(writer["price"]).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,'), '12C');
    sheet.setCellWithStyle('N18', '¥' + String(writer["sum"]).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,'), '12C');
    sheet.setCellWithStyle('S18', '', '12C');
    for (i = _l = 18; _l <= 30; i = ++_l) {
      if (i !== 30) {
        sheet.mergeCells([i - 1, 2], [i - 1, 7]);
      } else {
        sheet.mergeCells([i - 1, 1], [i - 1, 7]);
      }
      sheet.mergeCells([i - 1, 8], [i - 1, 9]);
      sheet.mergeCells([i - 1, 10], [i - 1, 12]);
      sheet.mergeCells([i - 1, 13], [i - 1, 17]);
      sheet.mergeCells([i - 1, 18], [i - 1, 20]);
    }
    sheet.setCellWithStyle('C28', '小計', '12C');
    sheet.setCellWithStyle('C29', writer["tax_name"] + '(' + writer["tax_val"] + '%)', '12C');
    sheet.setCellWithStyle('N28', '¥' + String(writer["sum"]).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,'), '12C');
    sheet.setCellWithStyle('N29', '¥-' + String(Math.ceil(writer["sum"] * writer["tax_val"] / 100)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,'), '12C');
    sheet.setCellWithStyle('B30', '合計', '12C');
    return sheet.setCellWithStyle('N30', '¥' + String(writer["sum"] - Math.ceil(writer["sum"] * writer["tax_val"] / 100)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,'), '12C');
  };

}).call(this);
