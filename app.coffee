express = require "express"
busboy = require "connect-busboy"
formidable = require "formidable"
util = require "util"
path = require "path"
app = express()
app.set('view engine', 'jade')
app.set 'views', __dirname + '/views'
app.use busboy()
# app.use express.static(path.join(__dirname, 'files'))
# express.logger('dev')

fs = require "fs"
officegen = require "officegen"
xlsx = officegen 'xlsx'

SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

OUTPUT_PATH = "ライター個人請求書.xlsx"

WRITER_COLUMN = 3    # 名前
PRICE_COLUMN  = 4    # 単価
NUM_COLUMN    = 5    # 本数
ID_COLUMN     = 6    # 請求番号
BANK1_COLUMN  = 7    # 銀行名
BANK2_COLUMN  = 8    # 支店名
BANK3_COLUMN  = 9    # 口座番号
BANK4_COLUMN  = 10   # 名義
ADDR1_COLUMN  = 11   # 郵便番号
ADDR2_COLUMN  = 12   # 住所①
ADDR3_COLUMN  = 13   # 住所②



COLUMNS_WIDTH = 2.2  # カラムの大きさ

xlsx.on 'finalize', (written) ->
	console.log 'Finish to create an Excel File. Total bytes created: ' + written

xlsx.on 'error', (err) ->
	console.log "Xlsx Err: " + err

app.get '/', (req, res) ->
	res.render 'index'

app.post '/upload', (req, res) ->
	form = new formidable.IncomingForm();
	form.parse req, (err, fields, files) ->
		old_path = files.file.path
		file_size = files.file.size
		file_ext = files.file.name.split('.').pop()
		index = old_path.lastIndexOf('/') + 1
		file_name = old_path.substr index
		new_path = path.join process.env.PWD, '/files/', file_name + '.' + file_ext

		fs.readFile old_path, (err, data) ->
			fs.writeFile new_path, data, (err) ->
				fs.unlink old_path, (err) ->
					if err
						res.status 500
						res.json 
							'success': false
					else
						res.status 200
						readFile new_path
						fs.readFile OUTPUT_PATH, (err, data) ->
							res.writeHead 200, 
								'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
							res.end data, 'binary'

	# console.log "aaa"
	# if req.files
	# 	console.log util.inspect(req.files)
	# fstream = undefined
 #  req.pipe req.busboy
 #  req.busboy.on 'file', (fieldname, file, filename) ->
 #    console.log "Uploading: " + filename 
 #    fstream = fs.createWriteStream __dirname + '/files/' + filename
 #    file.pipe fstream
 #    fstream.on 'close', () ->
 #    	console.log "upload finished"
	# 		readFile __dirname + '/files/' + filename

app.listen(3000)

readFile = (filePath) ->

	SpreadsheetReader.read filePath, (err, workbook) ->
		if err
			console.log err
			return
		writers = []
		workbook.sheets.forEach (sheet) ->
			console.log 'sheet: %s(%d)', sheet.name, sheet.index
			if sheet.index != 7
				return
			stopLoop = false
			sheet.rows.forEach (row) ->
				row.forEach (cell) ->
					if cell.row >= 4 && !stopLoop
						writer = {}
						if writers.length > cell.row - 4
							writer = writers[cell.row-4]
						else
							writers.push writer
						switch(cell.column)
							when WRITER_COLUMN then writer["name"] = cell.value
							when PRICE_COLUMN  then writer["price"] = cell.value
							when NUM_COLUMN
								writer["num"] = cell.value
								if cell.value && cell.value > 0
									writer["sum"] = writer["price"] * writer["num"]
							when ID_COLUMN
								writer["id"] = cell.value if cell.value
							when BANK1_COLUMN
								writer["bank1"] = cell.value if cell.value
							when BANK2_COLUMN
								writer["bank2"] = cell.value if cell.value
							when BANK3_COLUMN
								writer["bank3"] = cell.value if cell.value
							when BANK4_COLUMN
								writer["bank4"] = cell.value if cell.value
							when ADDR1_COLUMN
								writer["addr1"] = cell.value if cell.value
							when ADDR2_COLUMN
								writer["addr2"] = cell.value if cell.value
							when ADDR3_COLUMN
								writer["addr3"] = cell.value if cell.value
							
						if cell.column >= WRITER_COLUMN && (!writer["name"] || writer["name"] == "合計")
							stopLoop = true
		writers.forEach (writer) ->
			if(writer["name"] && writer["num"] && writer["num"] > 0 && writer["sum"] && writer["sum"] > 0)
				makeBill(writer)
		out = fs.createWriteStream OUTPUT_PATH

		out.on 'error', (err) ->
			console.log "Error: " + err

		xlsx.generate out

makeBill = (writer) ->
	sheet = xlsx.makeNewSheet()
	sheet.name = writer["name"]
	sheet.columnsWidth = []
	for i in [0..21]
		sheet.columnsWidth[i] = COLUMNS_WIDTH

	sheet.setCellWithStyle 'B2',  '御　請　求　書', '20B'
	sheet.setCellWithStyle 'B4',  '株式会社 Ｄｏｎｕｔｓ 御中', '18BU'

	dd = new Date()
	yy = dd.getYear()
	yy += 1900 if yy < 2000
	yy = yy - 2000 + 12
	mon = dd.getMonth() + 1
	dd = dd.getDate()
	sheet.setCellWithStyle 'Q2',  '平成' + yy + '年' + mon + '月' + dd + '日', '14BU'
	sheet.mergeCells [1,16], [1,20]

	sheet.setCellWithStyle 'Q1',  '請求番号: ' + writer["id"], '11'
	sheet.mergeCells [0,16], [0,19]
	sheet.setCellWithStyle 'B6',  '下記のとおり御請求申し上げます', '12'

	sheet.setCellWithStyle 'Q7',  writer["addr1"], '12' if writer["addr1"]
	sheet.setCellWithStyle 'Q8',  writer["addr2"], '12' if writer["addr2"]
	sheet.setCellWithStyle 'Q9',  writer["addr3"], '12' if writer["addr3"]
	sheet.setCellWithStyle 'Q10', writer["name"], '12'

	sheet.setCellWithStyle 'B11', '振込先銀行', '14BU'
	sheet.mergeCells [10,1], [10,3]
	sheet.setCellWithStyle 'B13', '口座番号', '14BU'
	sheet.mergeCells [12,1], [12,3]
	sheet.setCellWithStyle 'B14', '名義', '14BU'
	sheet.mergeCells [13,1], [13,3]
	sheet.setCellWithStyle 'E11', writer["bank1"], '14BU' if writer["bank1"]
	sheet.mergeCells [10,4], [10,8]
	sheet.setCellWithStyle 'E12', writer["bank2"], '14BU' if writer["bank2"]
	sheet.mergeCells [11,4], [11,8]
	sheet.setCellWithStyle 'E13', writer["bank3"], '14BU' if writer["bank3"]
	sheet.mergeCells [12,4], [12,8]
	sheet.setCellWithStyle 'E14', writer["bank4"], '14BU' if writer["bank4"]
	sheet.mergeCells [13,4], [13,8]

	sheet.setCellWithStyle 'B16', '合計金額', '24C'
	sheet.mergeCells [15,1], [15,7]
	sheet.setCellWithStyle 'I16', '¥'+String( (writer["sum"] - Math.floor(writer["sum"]*10.21/100)) ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '24C'
	sheet.mergeCells [15,8], [15,11]
	sheet.setCellWithStyle 'B17', '摘要', '12C'
	sheet.mergeCells [16,1], [16,7]
	sheet.setCellWithStyle 'I17', '数量', '12C'
	sheet.mergeCells [16,8], [16,9]
	sheet.setCellWithStyle 'K17', '単価', '12C'
	sheet.mergeCells [16,10], [16,12]
	sheet.setCellWithStyle 'N17', '金額(税抜)', '12C'
	sheet.mergeCells [16,13], [16,17]
	sheet.setCellWithStyle 'S17', '備考', '12C'
	sheet.mergeCells [16,18], [16,20]

	sheet.setCellWithStyle 'C18', '原稿料(' + mon + '月分)', '12C'
	sheet.setCellWithStyle 'I18', '' + writer["num"], '12C'
	sheet.setCellWithStyle 'K18', '¥' + String( writer["price"] ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'N18', '¥' + String( writer["sum"] ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'S18', '', '12C'

	for i in [18..30]
		if i-17 <= 10
			sheet.setCellWithStyle 'B'+i, ''+(i-17), '12C'
		if i!=30
			sheet.mergeCells [i-1,2], [i-1,7]
		else
			sheet.mergeCells [i-1,1], [i-1,7]
		sheet.mergeCells [i-1,8], [i-1,9]
		sheet.mergeCells [i-1,10], [i-1,12]
		sheet.mergeCells [i-1,13], [i-1,17]
		sheet.mergeCells [i-1,18], [i-1,20]

	sheet.setCellWithStyle 'C28', '小計', '12C'
	sheet.setCellWithStyle 'C29', '源泉徴収税(10.21%)', '12C'
	sheet.setCellWithStyle 'N28', '¥'+String( writer["sum"] ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'N29', '¥-'+String( Math.floor(writer["sum"]*10.21/100) ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'B30', '合計', '12C'
	sheet.setCellWithStyle 'N30', '¥'+String( (writer["sum"] - Math.floor(writer["sum"]*10.21/100)) ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	