Option Explicit

' �ύX����
' 2022/08/27 �V�K�쐬

' ========== �ϐ��錾 START ==========
' �萔��` START
' ��؂蕶��
Const CONST_STR_SEP = ": "

' �萔��` END

' �擾����PC��{���̕ϐ���` Start
' PC�Ǘ��ԍ�
Dim strPcId
' ���[�J�[
Dim strMaker
' �^��
Dim strProductNo
' �V���A���ԍ�
Dim strSerialNo
' CPU
Dim strCpu
' ������
Dim strMemory
' HDD�e��
Dim strHddSize
Dim dblHddSize
' �v���C���X�g�[��OS
Dim strPreOs
' �擾����PC��{���̕ϐ���` End

' ���̑����
' �h���C����
Dim charDomain
' ���[�U�[��(���݃��O�C�����Ă��郆�[�U�[)
Dim charUserName
' UUID
Dim charUuid
' �o�̓��b�Z�[�W(LAN�A�_�v�^�[)
Dim strLanAdpt

' LAN�A�_�v�^�[�̌�
Dim intCntLan

' �o�͏��
Dim strOutputMessage

' �I�u�W�F�N�g�ϐ��̒�` Start
' WMI�I�u�W�F�N�g
Dim objWMIService
Dim objLocator
Dim objServer

' �A�C�e���p�ϐ�
Dim objItem
Dim colItems

' �R�}���h���C�����s�I�u�W�F�N�g Start
' �V�F���I�u�W�F�N�g
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
' ���s�p�I�u�W�F�N�g
Dim objExec
' �t�@�C���o�͗p�I�u�W�F�N�g
Dim objFso
Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")
' �e�L�X�g�t�@�C���p�I�u�W�F�N�g
Dim objTextFile
' �R�}���h���C�����s�I�u�W�F�N�g END
' �I�u�W�F�N�g�ϐ��̒�` END

Dim strOutputDir
Dim strOutputFile
Dim strOutputFull
Dim strYmd
Dim strError

' ========== �ϐ��錾 END ==========

' �o�͐��`�P START
' �o�̓t�H���_ �{vbs�Ɠ����ꏊ�ɕۑ�
strOutputDir = "."

' �ۑ�������̃t�H���_�ɂ���ꍇ
'strOutputDir = "C:\temp"

' �ۑ�����t�@�C���T�[�o�[/NAS�ɂ���ꍇ
'strOutputDir = "\\nasne\share1"

' ���t�擾
strYmd = Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2)
' �o�͐��`�P END

' PC�Ǘ����I�u�W�F�N�g�̎擾
' �Ώۍ��ځF
'   OS�EOS�T�[�r�X�p�b�N
'   �R���s���[�^���E�h���C�����E���[�U�[���E�������e��
'   �x���_�[�E�@�햼�E�V���A���i���o�[
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' OS�E�T�[�r�X�p�b�N�̎擾 START
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)

For Each objItem in colItems
	' OS
	strPreOs = objItem.Caption
Next
' OS�E�T�[�r�X�p�b�N�̎擾 End

' �R���s���[�^���E�h���C�����E���[�U�[���E�������e�ʂ̎擾 START
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)

For Each objItem in colItems
	'�R���s���[�^��
	strPcId = objItem.Name
	'�h���C����
	charDomain = objItem.Domain
	' ���[�U�[��
	charUserName = objItem.UserName
	' �������e��
	strMemory = fix(objItem.TotalPhysicalMemory /1024 /1024) & "MB"
Next
' �R���s���[�^���E�h���C�����E���[�U�[���E�������e�ʂ̎擾 END

' �x���_�[�E�@�햼�E�V���A���i���o�[ START
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)

For Each objItem in colItems
	' �x���_�[
	strMaker = objItem.Vendor
	' �@�햼
	strProductNo = objItem.Name
	' �V���A���i���o�[
	strSerialNo = objItem.IdentifyingNumber
Next
' �x���_�[�E�@�햼�E�V���A���i���o�[ END

' CPU START
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_Processor")

For Each objItem In colItems
    strCpu = strCpu & objItem.Name & " " & objItem.CurrentClockSpeed & "MHz"
Next
' CPU END

' UUID�擾 START
Set objWMIService = GetObject("winmgmts:\\.")
Set colItems = objWMIService.InstancesOf("Win32_ComputerSystemProduct")

For Each objItem In colItems
	charUuid = objItem.Uuid
Next
' UUID�擾 END

' �G���[���͎��̏��� START
On Error Resume Next

' HDD START
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set colItems = objWMIService.ExecQuery("SELECT Caption,Size FROM Win32_DiskDrive")

For Each objItem In colItems
	' HDD�T�C�Y������
	dblHddSize = -1
	
	' HDD�T�C�Y���擾
	dblHddSize = objItem.Size
	
	' GB�P�ʂŎ擾
	dblHddSize = dblHddSize / (1000 * 1000 * 1000)

	' ������O�ʂ��l�̌ܓ�
	dblHddSize = round(dblHddSize, 2)

	' GB�\�L��ǉ�
	if strHddSize <> "" then
		strHddSize = strHddSize & ","
	end if
	
	strHddSize = strHddSize & dblHddSize & "GB"
Next
' HDD END

' LAN�A�_�v�^�[���擾 START
' ���[�J���R���s���[�^�ɐڑ�����B
Set objLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
Set objServer = objLocator.ConnectServer
' �N�G���[������WQL�ɂĎw�肷��B
Set colItems = objServer.ExecQuery("Select * From Win32_NetworkAdapterConfiguration")

' LAN�A�_�v�^�[�̌���������
intCntLan = 0

For Each objItem In colItems
	If objItem.IPEnabled = True Then
		' LAN�A�_�v�^�[�̃J�E���g���C���N�������g
		intCntLan = intCntLan + 1
		
		' ���s�R�[�h�̒ǉ�
		If strLanAdpt <> "" Then
			strLanAdpt = strLanAdpt & vbCrLf
		end if
		
		' ���b�Z�[�W�̒ǉ�
		strLanAdpt = strLanAdpt & _
			"LanAdapter_" & intCntLan & CONST_STR_SEP & objItem.Description & vbCrLf & _
			"IPAddress_" & intCntLan & CONST_STR_SEP & objItem.IPAddress(0) & vbCrLf & _
			"IPSubnet_" & intCntLan & CONST_STR_SEP & objItem.IPSubnet(0) & vbCrLf & _
			"MacAddress_" & intCntLan & CONST_STR_SEP & objItem.MACAddress
	End If
Next
' LAN�A�_�v�^�[���擾 END

' �G���[���b�Z�[�W������ꍇ�͓o�^
If Err.Number <> 0 Then
	strError = "Err.Number:" & Err.Number & "+" & "Err.Description:" & Err.Description
End If

' �G���[�����N���A����B
Err.Clear

' �G���[���̓��b�Z�[�W�\��
On Error Goto 0
' �G���[���͎��̏��� END

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

' �o�͐��`�Q START
' �o�̓t�@�C��
strOutputFile = strPcId & "_" & "pcMainInfo" & "_" & strYmd & ".txt"

' �o�̓t���p�X�̐ݒ�
strOutputFull = strOutputDir & "\" & strOutputFile

' �o�͐��`�Q END

' �ۑ���ւ̏������� Start
Set objTextFile = objFso.CreateTextFile(strOutputFull, True)

' ��������
objTextFile.WriteLine strOutputMessage

' �t�@�C�� CLOSE
objTextFile.close
' �ۑ���ւ̏������� End

' �I�����b�Z�[�W�����[�U�[�ɕ\��
Wscript.Echo "PC����ۑ����܂����B" & vbCrLf & _
	"�ۑ���t�H���_" & CONST_STR_SEP & strOutputDir & vbCrLf & _
	"�t�@�C����" & CONST_STR_SEP & strOutputFile & vbCrLf & _
	"----------" & vbCrLf & _
	strOutputMessage & vbCrLf & _
	"----------"


