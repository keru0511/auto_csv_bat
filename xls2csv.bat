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
echo USAGE:%~n0 �}�X�^�[ID �K�v��󔒋�؂�
echo.
echo       ���̃t�@�C����z�u�����f�B���N�g���ɂ���.
echo       xls�t�@�C�����J���ACSV(�J���}��؂�)�`���ŕۑ�(SaveAs)���܂��B
echo.
echo  ��j./xls2csv.bat a b c d
echo  A��̋󔒍s���Ȃ���B,C,D����o�͂��܂��B
echo.
echo  ��ڂ̒l�͗�̒l���󔒂̍s�����O����܂�
echo.
echo  ��ڈȍ~�̒l�͒��o��������͂��܂��i���͏��Ƀ\�[�g����܂��j
echo.
goto :eof

rem ********************************************************************************
rem */
@end
//---------------------------------------------------------- �Z�b�g�A�b�v

var Args = WScript.Arguments;
var EXCEL = WScript.CreateObject("EXCEL.Application");
var SHELL = WScript.CreateObject("WScript.Shell");
var fso = WScript.CreateObject("Scripting.FileSystemObject");

//���O�v��
function echo(o){ WScript.Echo(o); }

// EXCEL�̒萔
var xlCSV = 6;

//---------------------------------------------------------- ��������
var sheet = null;
var infile = null;
var outfile = null;

// �p�����[�^�[�̎擾
// �����f�[�^�̎擾���@��������Ȃ������̂łƂ肠�����R�}���h���C������擾
var key = Args(0);
var p = [];
for (var i = 1; i < Args.Length; i++){
	p[i] = conversion(Args(i));
}

// �t�@�C������TXT����擾
// �擾�����t�@�C������z��Ɋi�[����
// �t�@�C�������珇�Ԃɓǂݍ���
var file = fso.OpenTextFile("xls2csvconf.txt");
var txt = [];
var i = 0;
while (!file.AtEndOfStream) {
    var line = file.ReadLine();
	txt[i] = line;
	i++;
}
// �t�@�C����
var filenum = i;

//---------------------------------------------------------- �又��
// �����擾
function length(d){
	var used = d.UsedRange;
	if (used.Count <= 1) return; // �g�p���Z����1�ȉ��Ȃ珈�����Ȃ�
	var last = new Object();
	last.row = used.Cells(used.Count).Row;
	last.col = used.Cells(used.Count).Column;
	return last;
}
// �p���A���l�ϊ�
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

// �J�����g�f�B���N�g���̐؂�ւ�
if (EXCEL.DefaultFilePath != SHELL.CurrentDirectory){
	EXCEL.DefaultFilePath = SHELL.CurrentDirectory;
	delete EXCEL;
	EXCEL = WScript.CreateObject("EXCEL.Application");
}

//�V�K�u�b�N
var NB = EXCEL.Workbooks.Add();
EXCEL.DisplayAlerts = false;
var NS = NB.Worksheets(1);
var swrite = 1;

// �t�@�C�����J��
try{
	for(var n = 0; n < filenum; n++){
		echo(txt[n] + "���������Ă��܂��B" + "[" + n + '/' + filenum + "]");
		// �G�N�Z���̎擾
		var WB = EXCEL.Workbooks.Open(txt[n]);
		var WS = WB.Worksheets(1);
		var WBlength = length(WS);
		// ��������
		// ���o���ʒu���Ⴄ�ꍇ��j�̏����l��ς���
		for(var j = 4; j <= WBlength.row; j++){
			// �󔒂݂̗̂�𔻒�
			var ret = [];
			for(var i = 1; i <= WBlength.col; i++){
				ret[i] =  WS.Cells(j,i).value;
				var ID = WS.Cells(j,key).value;
			}
			// �}�X�^�[ID��̋󔒂̂��̂����O����
			// ��ɒl�����͂���Ă���ꍇ�̂ݏ�������
			if(ret.join('') != "" && ID !== undefined){
				for(var i = 1; i <= p.length; i++){
					NS.Cells(swrite,i) = ret[p[i]];
				}
				swrite++;
			}
		}
		WB.Close(false);
	}
	// xls�ϊ�
	NB.SaveAs("xls2csv.csv", xlCSV);
	// ����xls2csvconf.txt����΂�����ύX����
	fso.DeleteFolder("xls2csvconf.txt")
} catch(e){
	echo(e.number + ":" + e.description);
} finally {
	NB.Close(false);
	EXCEL.quit();
	echo('�������I�����܂��B')
}