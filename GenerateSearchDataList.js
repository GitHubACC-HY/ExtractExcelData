// �G�N�Z���t�@�C���������Ă���f�B���N�g���i���̃X�N���v�g����̑��΃p�X��z��j
var TargetDir = "./";

// �t�@�C���̊g���q�i�]�v�ȃt�@�C���ǂݍ��܂Ȃ��΍�j
var TargetFileExt = "xlsx";

// �Ώۂ̃V�[�g��
var TargetSheetName = "����";

// �e�[�u���̊J�n�s�i�w�b�_�ł͂Ȃ��f�[�^�̊J�n�j
var StartRowSpecifiedMode = 2; // 1: �Œ�l(TargetStartRow)�A2:����̒l���o��܂�A��𑖍�����
var TargetStartRow = 2;
var StartRowSearchStr = "No."

// �擾�Ώۃf�[�^
// �f�[�^�𐔒l�Ŏw�肷����̂ɂ��ẮAA��Ȃ�1�AB��Ȃ�2�A�A�A�̂悤�Ɏw�肷��B�t�@�C�����i�g���q�����j���g�������ꍇ��0���w�肷��
var KeyCols = [6, 7]; // �����L�[�Ƃ������f�[�^�̗�ꗗ�𐔒l�Ŏw��A�����w��̏ꍇ�͍����珇�ɊK�w�ɂȂ�B��Z�����܂ޏꍇ���̍s�𖳎�
var DataCols = [0, 2]; // �������ʂɏo�������f�[�^�̗�ꗗ�𐔒l�Ŏw��A�S�ċ�Z���̏ꍇ���̍s�𖳎�












var fs = WScript.CreateObject("Scripting.FileSystemObject");
main();
fs = null;

function main() {
	var data = new Object();
	
	// �w��K�w�ɂ���t�@�C���̗�
	var path = fs.GetAbsolutePathName(TargetDir);
	var files = fs.GetFolder(path).Files;
	var e = new Enumerator(files);
	for ( ; !e.atEnd(); e.moveNext()) {
		var file = e.item();
		// �g���q�Ńt�@�C�����t�B���^�����O
		if(fs.GetExtensionName(file.Path) == TargetFileExt) {
			LoadData(file.Path, data);
		}
	}
	
	// ��x�z��ō�����f�[�^�𕶎���ɕϊ�����i����������ǐ��̂��߂��̂܂܂Ƃ���B�j
	ArrayToStringRecursive(data);
	
	// �f�[�^��HTML�ň�����悤�AJson�`���ɕϊ�����
	var html = new ActiveXObject('htmlfile');
	html.write('<meta http-equiv="x-ua-compatible" content="IE=11" />');
	var JSON = html.parentWindow.JSON;
	var ret = JSON.stringify(data);
	
	// �t�@�C����������
	var file = fs.OpenTextFile( "searchDataList.js", 2, true, -2 );
	file.Write("var SearchData = " + ret + ";"); // Javascript���爵����悤�A�ϐ��錾�Ƃ���B

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
	
	var lastRow = sheet.UsedRange.Cells(sheet.UsedRange.Count).Row + 1; // ��肱�ڂ����|���̂�+1���Ă����i���t�@�����X��������Ȃ��E�E�E�j
	
	// �J�n�s�����߂�
	var startRow = -1;
	if(StartRowSpecifiedMode == 1) {
		startRow = TargetStartRow;
	}
	else if(StartRowSpecifiedMode == 2) {
		for(var i = 1; i < lastRow; i++) {
			if(sheet.Cells(i, 1).Value == StartRowSearchStr) {
				TargetStartRow = i+1;
			}
		}
	}
	
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

// Object���ċA�I�Ɍ������AArray��Json������ɕϊ�����
function ArrayToStringRecursive(data) {
	for(var key in data) {
		if(!data[key].length && !data[key].push) {
			ArrayToStringRecursive(data[key]);
		} else {
			data[key] = ArrayToString(data[key]);
		}
	}
}

// Array��Json������ɕϊ�����
function ArrayToString(array) {
	var ret = "[";
	for(var i = 0; i < array.length; i++) {
		if(i != 0) ret += ", ";
		ret += array[i];
	}
	return ret + "]";
}

