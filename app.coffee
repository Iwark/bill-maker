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
xlsx = undefined

SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader;

OUTPUT_PATH = "writer-bill.xlsx"

WRITER_COLUMN = 3    # 名前
TAX_COLUMN    = 4    # 個人（源泉徴収税） / 法人（消費税）
PRICE_COLUMN  = 5    # 単価
NUM_COLUMN    = 6    # 本数
ID_COLUMN     = 7    # 請求番号
BANK1_COLUMN  = 8    # 銀行名
BANK2_COLUMN  = 9    # 支店名
BANK3_COLUMN  = 10   # 口座番号
BANK4_COLUMN  = 11   # 名義
ADDR1_COLUMN  = 12   # 郵便番号
ADDR2_COLUMN  = 13   # 住所①
ADDR3_COLUMN  = 14   # 住所②

COLUMNS_WIDTH = 2.2  # カラムの大きさ

app.get '/', (req, res) ->
	res.render 'index'

app.post '/upload', (req, res) ->
	xlsx = officegen 'xlsx'
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
							when TAX_COLUMN
								if cell.value == "個人"
									writer["tax_name"] = "源泉徴収税"
									writer["tax_val"] = 10.21
								else if cell.value == "法人"
									writer["tax_name"] = "消費税"
									writer["tax_val"] = 8.00
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

		xlsx.generate out,
			finalize: (written) ->
				console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' )
			error: (err) ->
				console.log ( err )

makeBill = (writer) ->
	sheet = xlsx.makeNewSheet()
	sheet.name = writer["name"]
	sheet.columnsWidth = []
	for i in [0..21]
		sheet.columnsWidth[i] = COLUMNS_WIDTH

	sheet.setCellWithStyle 'B2',  '御　請　求　書', '20B'
	sheet.mergeCells [3,1], [3,6]
	sheet.setCellWithStyle 'B4',  '株式会社 Ｄｏｎｕｔｓ 御中', '18BU'
	sheet.setCellWithStyle 'C4', 'a', '18BU'
	sheet.setCellWithStyle 'D4', 'a', '18BU'
	sheet.setCellWithStyle 'E4', 'a', '18BU'
	sheet.setCellWithStyle 'F4', 'a', '18BU'
	sheet.setCellWithStyle 'G4', 'a', '18BU'

	dd = new Date()
	yy = dd.getYear()
	yy += 1900 if yy < 2000
	yy = yy - 2000 + 12
	mon = dd.getMonth() + 1
	dd = dd.getDate()
	sheet.mergeCells [1,16], [1,20]
	sheet.setCellWithStyle 'Q2',  '平成' + yy + '年' + mon + '月' + dd + '日', '14BU'
	sheet.setCellWithStyle 'R2', 'a', '14BU'
	sheet.setCellWithStyle 'S2', 'a', '14BU'
	sheet.setCellWithStyle 'T2', 'a', '14BU'

	sheet.mergeCells [0,16], [0,19]
	sheet.setCellWithStyle 'Q1',  '請求番号: ' + writer["id"], '11' if writer["id"]
	sheet.setCellWithStyle 'B6',  '下記のとおり御請求申し上げます', '12'

	sheet.setCellWithStyle 'Q7',  writer["addr1"], '12' if writer["addr1"]
	sheet.setCellWithStyle 'Q8',  writer["addr2"], '12' if writer["addr2"]
	sheet.setCellWithStyle 'Q9',  writer["addr3"], '12' if writer["addr3"]
	sheet.setCellWithStyle 'Q10', writer["name"], '12'

	sheet.mergeCells [10,1], [10,3]
	sheet.setCellWithStyle 'B11', '振込先銀行', '14BU'
	sheet.setCellWithStyle 'C11', 'a', '14BU'
	sheet.setCellWithStyle 'D11', 'a', '14BU'
	sheet.mergeCells [12,1], [12,3]
	sheet.setCellWithStyle 'B13', '口座番号', '14BU'
	sheet.setCellWithStyle 'C13', 'a', '14BU'
	sheet.setCellWithStyle 'D13', 'a', '14BU'
	sheet.mergeCells [13,1], [13,3]
	sheet.setCellWithStyle 'B14', '名義', '14BU'
	sheet.setCellWithStyle 'C14', 'a', '14BU'
	sheet.setCellWithStyle 'D14', 'a', '14BU'
	sheet.mergeCells [10,4], [10,8]
	sheet.setCellWithStyle 'E11', writer["bank1"], '14BU' if writer["bank1"]
	sheet.setCellWithStyle 'F11', 'a', '14BU'
	sheet.setCellWithStyle 'G11', 'a', '14BU'
	sheet.setCellWithStyle 'H11', 'a', '14BU'
	sheet.setCellWithStyle 'I11', 'a', '14BU'
	sheet.mergeCells [11,4], [11,8]
	sheet.setCellWithStyle 'E12', writer["bank2"], '14BU' if writer["bank2"]
	sheet.setCellWithStyle 'F12', 'a', '14BU'
	sheet.setCellWithStyle 'G12', 'a', '14BU'
	sheet.setCellWithStyle 'H12', 'a', '14BU'
	sheet.setCellWithStyle 'I12', 'a', '14BU'
	sheet.mergeCells [12,4], [12,8]
	sheet.setCellWithStyle 'E13', writer["bank3"], '14BU' if writer["bank3"]
	sheet.setCellWithStyle 'F13', 'a', '14BU'
	sheet.setCellWithStyle 'G13', 'a', '14BU'
	sheet.setCellWithStyle 'H13', 'a', '14BU'
	sheet.setCellWithStyle 'I13', 'a', '14BU'
	sheet.mergeCells [13,4], [13,8]
	sheet.setCellWithStyle 'E14', writer["bank4"], '14BU' if writer["bank4"]
	sheet.setCellWithStyle 'F14', 'a', '14BU'
	sheet.setCellWithStyle 'G14', 'a', '14BU'
	sheet.setCellWithStyle 'H14', 'a', '14BU'
	sheet.setCellWithStyle 'I14', 'a', '14BU'

	sheet.mergeCells [15,1], [15,7]
	sheet.setCellWithStyle 'B16', '合計金額', '24C'
	sheet.setCellWithStyle 'C16', 'a', '24C'
	sheet.setCellWithStyle 'D16', 'a', '24C'
	sheet.setCellWithStyle 'E16', 'a', '24C'
	sheet.setCellWithStyle 'F16', 'a', '24C'
	sheet.setCellWithStyle 'G16', 'a', '24C'
	sheet.setCellWithStyle 'H16', 'a', '24C'
	sheet.mergeCells [15,8], [15,11]
	sheet.setCellWithStyle 'I16', '¥'+String( (writer["sum"] - Math.ceil(writer["sum"]*writer["tax_val"]/100)) ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '22C'
	sheet.setCellWithStyle 'J16', 'a', '22C'
	sheet.setCellWithStyle 'K16', 'a', '22C'
	sheet.setCellWithStyle 'L16', 'a', '22C'
	sheet.setCellWithStyle 'M16', ' ', 'mdB'
	sheet.setCellWithStyle 'N16', ' ', 'mdB'
	sheet.setCellWithStyle 'O16', ' ', 'mdB'
	sheet.setCellWithStyle 'P16', ' ', 'mdB'
	sheet.setCellWithStyle 'Q16', ' ', 'mdB'
	sheet.setCellWithStyle 'R16', ' ', 'mdB'
	sheet.setCellWithStyle 'S16', ' ', 'mdB'
	sheet.setCellWithStyle 'T16', ' ', 'mdB'
	sheet.setCellWithStyle 'U16', ' ', 'mdB'

	for i in [17..30]
		for j in [1..20]
			sheet.setCellWithStyle String.fromCharCode('A'.charCodeAt(0)+j) + i, ' ', 'thB'
			if j == 1 || j==7 || j==9 || j==12 || j==17
				if j!=1 || i!=30 
					if j==1 && i-17 <= 10
						sheet.setCellWithStyle String.fromCharCode('A'.charCodeAt(0)+j) + i, ''+(i-17), 'thBR'
					else
						sheet.setCellWithStyle String.fromCharCode('A'.charCodeAt(0)+j) + i, ' ', 'thBR'

	sheet.mergeCells [16,1], [16,7]
	sheet.setCellWithStyle 'B17', '摘要', '12C'
	sheet.mergeCells [16,8], [16,9]
	sheet.setCellWithStyle 'I17', '数量', '12C'
	sheet.mergeCells [16,10], [16,12]
	sheet.setCellWithStyle 'K17', '単価', '12C'
	sheet.mergeCells [16,13], [16,17]
	sheet.setCellWithStyle 'N17', '金額(税抜)', '12C'
	sheet.mergeCells [16,18], [16,20]
	sheet.setCellWithStyle 'S17', '備考', '12C'

	sheet.setCellWithStyle 'C18', '原稿料(' + mon + '月分)', '12C'
	sheet.setCellWithStyle 'I18', '' + writer["num"], '12C'
	sheet.setCellWithStyle 'K18', '¥' + String( writer["price"] ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'N18', '¥' + String( writer["sum"] ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'S18', '', '12C'

	for i in [18..30]
		if i!=30
			sheet.mergeCells [i-1,2], [i-1,7]
		else
			sheet.mergeCells [i-1,1], [i-1,7]
		sheet.mergeCells [i-1,8], [i-1,9]
		sheet.mergeCells [i-1,10], [i-1,12]
		sheet.mergeCells [i-1,13], [i-1,17]
		sheet.mergeCells [i-1,18], [i-1,20]
	sheet.setCellWithStyle 'C28', '小計', '12C'
	sheet.setCellWithStyle 'C29', writer["tax_name"] + '(' + writer["tax_val"] + '%)', '12C'
	sheet.setCellWithStyle 'N28', '¥'+String( writer["sum"] ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'N29', '¥-'+String( Math.ceil(writer["sum"]*writer["tax_val"]/100) ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
	sheet.setCellWithStyle 'B30', '合計', '12C'
	sheet.setCellWithStyle 'N30', '¥'+String( (writer["sum"] - Math.ceil(writer["sum"]*writer["tax_val"]/100)) ).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,' ), '12C'
