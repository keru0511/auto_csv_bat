@if (1==1) /*
@echo off

dir *.xlsx /B > xls2csvconf.txt

if "%~2"=="" goto :USAGE
if "%~1"=="/?" goto :USAGE

rem ********************************************************************************
:MAIN
CScript //nologo //E:JScript "%~f0" %*
If ERRORLEVEL 1 goto :USAGE
del xls2csvconf.txt
goto :eof

rem ********************************************************************************
:USAGE
echo USAGE:%~n0 マスターID 必要列空白区切り
echo.
echo       このファイルを配置したディレクトリにある.
echo       xlsファイルを開き、CSV(カンマ区切り)形式で保存(SaveAs)します。
echo.
echo  例）./xls2csv.bat a b c d
echo  A列の空白行を省いたB,C,D列を出力します。
echo.
echo  一つ目の値は列の値が空白の行が除外されます
echo.
echo  二つ目以降の値は抽出する列を入力します（入力順にソートされます）
echo.
goto :eof

rem ********************************************************************************
rem */
@end
//---------------------------------------------------------- セットアップ

var Args = WScript.Arguments;
var EXCEL = WScript.CreateObject("EXCEL.Application");
var SHELL = WScript.CreateObject("WScript.Shell");
var fso = WScript.CreateObject("Scripting.FileSystemObject");

//ログ要因
function echo(o){ WScript.Echo(o); }

// EXCELの定数
var xlCSV = 6;

//---------------------------------------------------------- 引数処理
var sheet = null;
var infile = null;
var outfile = null;

// パラメーターの取得
// いいデータの取得方法が見つからなかったのでとりあえずコマンドラインから取得
var key = Args(0);
var p = [];
for (var i = 1; i < Args.Length; i++){
	p[i] = conversion(Args(i));
}

// ファイル名をTXTから取得
// 取得したファイル名を配列に格納する
// ファイル名から順番に読み込む
var file = fso.OpenTextFile("xls2csvconf.txt");
var txt = [];
var i = 0;
while (!file.AtEndOfStream) {
    var line = file.ReadLine();
	txt[i] = line;
	i++;
}
// ファイル数
var filenum = i;

//---------------------------------------------------------- 主処理
// 長さ取得
function length(d){
	var used = d.UsedRange;
	if (used.Count <= 1) return; // 使用中セルが1以下なら処理しない
	var last = new Object();
	last.row = used.Cells(used.Count).Row;
	last.col = used.Cells(used.Count).Column;
	return last;
}
// 英字、数値変換
function conversion(p){
	var input = p;
	if (!input) return;

	var symbols = 'abcdefghijklmnopqrstuvwxyz';
	var result;
	if (/^[0-9]+$/.test(input)) {
		input = parseInt(input);
		result = [];
		while (input > 0) {
			result.unshift(symbols.charAt((input - 1) % symbols.length));
			input = Math.floor((input - 1) / symbols.length);
		}
		result = result.join('').toUpperCase();
	}else{
		result = 0;
		input = input.toLowerCase().split('').reverse();
		for (var i = 0, maxi = input.length; i < maxi; i++) {
			result += (symbols.indexOf(input[i]) + 1) * Math.pow(symbols.length, i);
		}
	}
	return result;
}

// カレントディレクトリの切り替え
if (EXCEL.DefaultFilePath != SHELL.CurrentDirectory){
	EXCEL.DefaultFilePath = SHELL.CurrentDirectory;
	delete EXCEL;
	EXCEL = WScript.CreateObject("EXCEL.Application");
}

//新規ブック
var NB = EXCEL.Workbooks.Add();
EXCEL.DisplayAlerts = false;
var NS = NB.Worksheets(1);
var swrite = 1;

// ファイルを開く
try{
	for(var n = 0; n < filenum; n++){
		echo(txt[n] + "を処理しています。" + "[" + n + '/' + filenum + "]");
		// エクセルの取得
		var WB = EXCEL.Workbooks.Open(txt[n]);
		var WS = WB.Worksheets(1);
		var WBlength = length(WS);
		// 書き込み
		// 見出し位置が違う場合はjの初期値を変える
		for(var j = 4; j <= WBlength.row; j++){
			// 空白のみの列を判定
			var ret = [];
			for(var i = 1; i <= WBlength.col; i++){
				ret[i] =  WS.Cells(j,i).value;
				var ID = WS.Cells(j,key).value;
			}
			// マスターID列の空白のものを除外する
			// 列に値が入力されている場合のみ書き込む
			if(ret.join('') != "" && ID !== undefined){
				for(var i = 1; i <= p.length; i++){
					NS.Cells(swrite,i) = ret[p[i]];
				}
				swrite++;
			}
		}
		WB.Close(false);
	}
	// xls変換
	NB.SaveAs("xls2csv.csv", xlCSV);
	// 他でxls2csvconf.txtあればここを変更する
	fso.DeleteFolder("xls2csvconf.txt")
} catch(e){
	echo(e.number + ":" + e.description);
} finally {
	NB.Close(false);
	EXCEL.quit();
	echo('処理を終了します。')
}