Option Explicit

' 変更履歴
' 2022/08/27 新規作成

' ========== 変数宣言 START ==========
' 定数定義 START
' 区切り文字
Const CONST_STR_SEP = ": "

' 定数定義 END

' 取得するPC基本情報の変数定義 Start
' PC管理番号
Dim strPcId
' メーカー
Dim strMaker
' 型番
Dim strProductNo
' シリアル番号
Dim strSerialNo
' CPU
Dim strCpu
' メモリ
Dim strMemory
' HDD容量
Dim strHddSize
Dim dblHddSize
' プレインストールOS
Dim strPreOs
' 取得するPC基本情報の変数定義 End

' その他情報
' ドメイン名
Dim charDomain
' ユーザー名(現在ログインしているユーザー)
Dim charUserName
' UUID
Dim charUuid
' 出力メッセージ(LANアダプター)
Dim strLanAdpt

' LANアダプターの個数
Dim intCntLan

' 出力情報
Dim strOutputMessage

' オブジェクト変数の定義 Start
' WMIオブジェクト
Dim objWMIService
Dim objLocator
Dim objServer

' アイテム用変数
Dim objItem
Dim colItems

' コマンドライン実行オブジェクト Start
' シェルオブジェクト
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
' 実行用オブジェクト
Dim objExec
' コマンドライン実行オブジェクト END
' オブジェクト変数の定義 END

Dim strMsg
Dim strOutputDir
Dim strOutputFile
Dim strOutputFull
Dim strYmd
Dim strError

Dim fso
Dim textFile
' ========== 変数宣言 END ==========

' 出力先定義１ START
' 出力フォルダ 本vbsと同じ場所に保存
strOutputDir = "."

' 保存先を特定のフォルダにする場合
'strOutputDir = "C:\temp"

' 保存先をファイルサーバー/NASにする場合
'strOutputDir = "\\nasne\share1"

' 日付取得
strYmd = Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2)
' 出力先定義１ END

' PC管理情報オブジェクトの取得
' 対象項目：
'   OS・OSサービスパック
'   コンピュータ名・ドメイン名・ユーザー名・メモリ容量
'   ベンダー・機種名・シリアルナンバー
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' OS・サービスパックの取得 START
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)

For Each objItem in colItems
	' OS
	strPreOs = objItem.Caption
Next
' OS・サービスパックの取得 End

' コンピュータ名・ドメイン名・ユーザー名・メモリ容量の取得 START
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)

For Each objItem in colItems
	'コンピュータ名
	strPcId = objItem.Name
	'ドメイン名
	charDomain = objItem.Domain
	' ユーザー名
	charUserName = objItem.UserName
	' メモリ容量
	strMemory = fix(objItem.TotalPhysicalMemory /1024 /1024) & "MB"
Next
' コンピュータ名・ドメイン名・ユーザー名・メモリ容量の取得 END

' ベンダー・機種名・シリアルナンバー START
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)

For Each objItem in colItems
	' ベンダー
	strMaker = objItem.Vendor
	' 機種名
	strProductNo = objItem.Name
	' シリアルナンバー
	strSerialNo = objItem.IdentifyingNumber
Next
' ベンダー・機種名・シリアルナンバー END

' CPU START
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_Processor")

For Each objItem In colItems
    strCpu = strCpu & objItem.Name & " " & objItem.CurrentClockSpeed & "MHz"
Next
' CPU END

' UUID取得 START
Set objWMIService = GetObject("winmgmts:\\.")
Set colItems = objWMIService.InstancesOf("Win32_ComputerSystemProduct")

For Each objItem In colItems
	charUuid = objItem.Uuid
Next
' UUID取得 END

' エラー時は次の処理 START
On Error Resume Next

' HDD START
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set colItems = objWMIService.ExecQuery("SELECT Caption,Size FROM Win32_DiskDrive")

For Each objItem In colItems
	' HDDサイズ初期化
	dblHddSize = -1
	
	' HDDサイズを取得
	dblHddSize = objItem.Size
	
	' GB単位で取得
	dblHddSize = dblHddSize / (1000 * 1000 * 1000)

	' 小数第三位を四捨五入
	dblHddSize = round(dblHddSize, 2)

	' GB表記を追加
	if strHddSize <> "" then
		strHddSize = strHddSize & ","
	end if
	
	strHddSize = strHddSize & dblHddSize & "GB"
Next
' HDD END

' LANアダプター情報取得 START
' ローカルコンピュータに接続する。
Set objLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set objServer = objLocator.ConnectServer
' クエリー条件をWQLにて指定する。
Set colItems = objServer.ExecQuery("Select * From Win32_NetworkAdapterConfiguration")

' LANアダプターの個数を初期化
intCntLan = 0

For Each objItem In colItems
	If objItem.IPEnabled = True Then
		' LANアダプターのカウントをインクリメント
		intCntLan = intCntLan + 1
		
		' 改行コードの追加
		If strLanAdpt <> "" Then
			strLanAdpt = strLanAdpt & vbCrLf
		end if
		
		' メッセージの追加
		strLanAdpt = strLanAdpt & _
			"LanAdapter_" & intCntLan & CONST_STR_SEP & objItem.Description & vbCrLf & _
			"IPAddress_" & intCntLan & CONST_STR_SEP & objItem.IPAddress(0) & vbCrLf & _
			"IPSubnet_" & intCntLan & CONST_STR_SEP & objItem.IPSubnet(0) & vbCrLf & _
			"MacAddress_" & intCntLan & CONST_STR_SEP & objItem.MACAddress
	End If
Next
' LANアダプター情報取得 END

' エラーメッセージがある場合は登録
If Err.Number <> 0 Then
	strError = "Err.Number:" & Err.Number & "+" & "Err.Description:" & Err.Description
End If

' エラー情報をクリアする。
Err.Clear

' エラー時はメッセージ表示
On Error Goto 0
' エラー時は次の処理 END

strOutputMessage = _
	"pcId" & CONST_STR_SEP & strPcId & vbCrLf & _
	"productNo" & CONST_STR_SEP & strProductNo & vbCrLf & _
	"serialNo" & CONST_STR_SEP & strSerialNo & vbCrLf & _
	"cpu" & CONST_STR_SEP & strCpu & vbCrLf & _
	"hddSize" & CONST_STR_SEP & strHddSize & vbCrLf & _
	"maker" & CONST_STR_SEP & strMaker & vbCrLf & _
	"memory" & CONST_STR_SEP & strMemory & vbCrLf & _
	"preOs" & CONST_STR_SEP & strPreOs & vbCrLf & _
	"charDomain" & CONST_STR_SEP & charDomain & vbCrLf & _
	"charUserName" & CONST_STR_SEP & charUserName & vbCrLf & _
	"charUuid" & CONST_STR_SEP & charUuid & vbCrLf & _
	strLanAdpt & vbCrLf & _
	"errorMessage" & CONST_STR_SEP & strError

' 出力先定義２ START
' 出力ファイル
strOutputFile = strPcId & "_" & "pcMainInfo" & "_" & strYmd & ".txt"

' 出力フルパスの設定
strOutputFull = strOutputDir & "\" & strOutputFile

' 出力先定義２ END

' 保存先への書き込み Start
Set fso = Wscript.CreateObject("Scripting.FileSystemObject")
Set textFile = fso.CreateTextFile(strOutputFull, True)

' 書き込み
textFile.WriteLine strOutputMessage

' ファイル CLOSE
textFile.close
' 保存先への書き込み End

Wscript.Echo "PC情報を保存しました。" & vbCrLf & _
	"保存先フォルダ" & CONST_STR_SEP & strOutputDir & vbCrLf & _
	"ファイル名" & CONST_STR_SEP & strOutputFile & vbCrLf & _
	"----------" & vbCrLf & _
	strOutputMessage & vbCrLf & _
	"----------"

' 出力処理 End

