// エクセルファイルが入っているディレクトリ（このスクリプトからの相対パスを想定）
var TargetDir = "./";

// ファイルの拡張子（余計なファイル読み込まない対策）
var TargetFileExt = "xlsx";

// 対象のシート名
var TargetSheetName = "操作";

// テーブルの開始行（ヘッダではなくデータの開始）
var TargetStartRow = 2;

// 取得対象データ
// データを数値で指定するものについては、A列なら1、B列なら2、、、のように指定する。ファイル名（拡張子無し）を使いたい場合は0を指定する
var KeyCols = [6, 7]; // 検索キーとしたいデータの列一覧を数値で指定、複数指定の場合は左から順に階層になる。空セルを含む場合その行を無視
var DataCols = [0, 2]; // 検索結果に出したいデータの列一覧を数値で指定、全て空セルの場合その行を無視












var fs = WScript.CreateObject("Scripting.FileSystemObject");
main();
fs = null;

function main() {
	var data = new Object();
	
	// 指定階層にあるファイルの列挙
	var path = fs.GetAbsolutePathName(TargetDir);
	var files = fs.GetFolder(path).Files;
	var e = new Enumerator(files);
	for ( ; !e.atEnd(); e.moveNext()) {
		var file = e.item();
		// 拡張子でファイルをフィルタリング
		if(fs.GetExtensionName(file.Path) == TargetFileExt) {
			LoadData(file.Path, data);
		}
	}
	
	// 一度配列で作ったデータを文字列に変換する（非効率だが可読性のためこのままとする。）
	ArrayToStringRecursive(data);
	
	// データをHTMLで扱えるよう、Json形式に変換する
	var html = new ActiveXObject('htmlfile');
	html.write('<meta http-equiv="x-ua-compatible" content="IE=11" />');
	var JSON = html.parentWindow.JSON;
	var ret = JSON.stringify(data);
	
	// ファイル書き込み
	var file = fs.OpenTextFile( "searchDataList.js", 2, true, -2 );
	file.Write("var SearchData = " + ret + ";"); // Javascriptから扱えるよう、変数宣言とする。

	file.Close();
}

function LoadData(filePath, result) {
	var excel = WScript.CreateObject("Excel.Application");
	var book = excel.Workbooks.Open(filePath, 0, true);
	var sheet = null;
	for (var i = 0; i < book.Worksheets.Count; i++) {
		var tmp = book.Worksheets(i + 1);
		if(tmp.Name == TargetSheetName) {
			sheet = tmp;
			break;
		}
	}
	if(!sheet) return;
	
	var lastRow = sheet.UsedRange.Cells(sheet.UsedRange.Count).Row + 1; // 取りこぼしが怖いので+1しておく（リファレンスが見つからない・・・）
	for(var i = TargetStartRow; i < lastRow; i++) {
		var currentData = new Array();
		var hasDefined = false;
		for(var j = 0; j < DataCols.length; j++) {
			var value;
			if(DataCols[j] == 0) value = fs.GetBaseName(filePath);
			else value = sheet.Cells(i, DataCols[j]).Value;
			if(!value){
				continue;
			}
			hasDefined = true;
			currentData[currentData.length] = '"' + value.replace(new RegExp("\n", "g"), "") + '"';
		}
		if(!hasDefined) continue;
		
		var target = result;
		var hasUndefined = false;
		var lastKeys = new Array();
		for(var j = 0; j < KeyCols.length; j++) {
			var val = sheet.Cells(i, KeyCols[j]).Value;
			if(!val) {
				hasUndefined = true;
				break;
			}
			if(j == KeyCols.length - 1) {
				lastKeys = val.split("\n");
				break;
			}
			
			if(!target[val]) {
				target[val] = new Object();
			}
			target = target[val];
		}
 		if(!hasUndefined) {
 			for(var j = 0; j < lastKeys.length; j++) {
 				if(!target[lastKeys[j]]) target[lastKeys[j]] = new Array();
 				target[lastKeys[j]][target[lastKeys[j]].length] = ArrayToString(currentData);
 			}
 		}
	}
	book.Close();
	excel.Quit()
}

// Objectを再帰的に検索し、ArrayをJson文字列に変換する
function ArrayToStringRecursive(data) {
	for(var key in data) {
		if(!data[key].length && !data[key].push) {
			ArrayToStringRecursive(data[key]);
		} else {
			data[key] = ArrayToString(data[key]);
		}
	}
}

// ArrayをJson文字列に変換する
function ArrayToString(array) {
	var ret = "[";
	for(var i = 0; i < array.length; i++) {
		if(i != 0) ret += ", ";
		ret += array[i];
	}
	return ret + "]";
}

