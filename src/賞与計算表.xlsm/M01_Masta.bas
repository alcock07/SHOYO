Attribute VB_Name = "M01_Masta"
Option Explicit

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER9 = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const MYSERVER = "Data Source=HB14\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD9 = "Password=ALCadmin!;"
Public Const PSWD = "Password=admin;"

Public Const dbM As String = "\\192.168.128.4\hb\kyuyo\�O���[�v����.accdb"
Public Const dbT As String = "\\192.168.128.4\hb\ta\���^�V�X�e��\�O���[�v����.accdb"
Public Const dbS As String = "\\192.168.128.4\hb\kyuyo\�ܗ^\�ܗ^�f�[�^.accdb"
Public strSQL As String
Public strDB  As String

Sub Set_KBN()

Dim strKBN As String
Dim Index  As Long

    strKBN = Range("O2")
    For Index = 3 To 8
        If Cells(Index, 15) = strKBN Then
            Range("P2") = Cells(Index, 16)
            Range("Q2") = Cells(Index, 17)
            Range("R2") = Cells(Index, 18)
            Exit For
        End If
    Next Index
    
End Sub

'============================
'���^�}�X�^�̏ܗ^�敪�ύX���
'============================

Sub Get_Masta()

Dim cnA    As New ADODB.Connection
Dim rsA    As New ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strKBN As String
Dim lngR   As Long
Dim lngC   As Long
Dim DateA  As Date
Dim DateB  As Date
Dim strYY  As String
Dim lngMM  As Long

    '��ď�����
    Range("A4:J152").ClearContents
    Range("L4:L52").ClearContents
    Range("N4:N52").ClearContents
        
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    '���Ə��敪���ƓǍ���
    strKBN = Range("Q2")
    strSQL = ""
    strSQL = strSQL & "SELECT SKBN"
    strSQL = strSQL & "       ,SCODE"
    strSQL = strSQL & "       ,SNAME"
    strSQL = strSQL & "       ,SKBN"
    strSQL = strSQL & "       ,CLASS"
    strSQL = strSQL & "       ,PAY1"
    strSQL = strSQL & "       ,PAY2"
    strSQL = strSQL & "       ,OPT1"  '�Ǘ��E�蓖
    strSQL = strSQL & "       ,OPT2"  '�Ƒ��蓖
    strSQL = strSQL & "       ,BMN1"
    strSQL = strSQL & "       ,BMN2"
    strSQL = strSQL & "       ,BMN3"
    strSQL = strSQL & "       ,BMNNM"
    strSQL = strSQL & "       ,DATE2" '���ДN����
    strSQL = strSQL & "     FROM KYUMTA"
    strSQL = strSQL & "        WHERE KBN ='" & strKBN & "'"
    strSQL = strSQL & "        AND DATKB ='1'"
    strSQL = strSQL & "     ORDER BY SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 4
    Do Until rsA.EOF
        '�e���ھ��
        For lngC = 0 To 8
            Cells(lngR, lngC + 1) = rsA.Fields(lngC)
        Next lngC
        '����敪���
        If IsNull(rsA.Fields("BMN2")) = False Then Cells(lngR, 10) = rsA.Fields("BMN2")
        If IsNull(rsA.Fields("BMN3")) = False Then Cells(lngR, 12) = rsA.Fields("BMN3")
        '�V���Ј����菈��
        If rsA.Fields("DATE2") <> "" Then
            DateA = rsA.Fields("DATE2")
        End If
        strYY = Format(Now(), "yyyy")
        lngMM = Format(Now(), "m")
        If lngMM >= 4 And lngMM <= 7 Then
            lngMM = 1
        ElseIf lngMM >= 10 And lngMM <= 12 Then
            lngMM = 5
        Else
            lngMM = 0
        End If
        If lngMM > 0 Then
            DateB = strYY & "/" & Format(lngMM, "00") & "/01"
            If DateA > DateB Then
                Cells(lngR, 14) = "��"
            Else
                Cells(lngR, 14) = ""
            End If
        End If
        rsA.MoveNext
        lngR = lngR + 1
    Loop
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub

Sub Up_Masta()

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim Cmd    As New ADODB.Command
Dim strKBN As String
Dim strCD  As String
Dim strKB1 As String
Dim strKB2 As String
Dim strKB3 As String
Dim lngR   As Long
Dim lngC   As Long
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    strKBN = Range("Q2")
    lngR = 4
    Do
        strCD = Cells(lngR, 2) '�Ј�����
        If strCD = "" Then Exit Do
        strKB1 = Range("P2")
        strKB2 = Cells(lngR, 10)
        strKB3 = Cells(lngR, 12)
        If strCD <> "" Then
            strSQL = ""
            strSQL = strSQL & "SELECT BMN1"
            strSQL = strSQL & "       ,BMN2"
            strSQL = strSQL & "       ,BMN3"
            strSQL = strSQL & "       ,BMNNM"
            strSQL = strSQL & "       ,YKBN"
            strSQL = strSQL & "     FROM KYUMTA"
            strSQL = strSQL & "        WHERE SCODE = '" & strCD & "'"
            rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If rsA.EOF = False Then
                rsA.MoveFirst
                rsA.Fields(0) = strKB1 '���_�敪
                rsA.Fields(1) = strKB2 '�ܗ^�敪
                rsA.Fields(2) = strKB3 '�����敪
                rsA.Fields(3) = Cells(lngR, 13) '������
                If Cells(lngR, 14) = "��" Then
                    rsA.Fields(4) = "Y" '�V���Ј��敪
                Else
                    rsA.Fields(4) = ""
                End If
            End If
            rsA.Update
            rsA.Close
        End If
        lngR = lngR + 1
    Loop
    
    MsgBox "�o�^���܂���(^^��", vbInformation, "�}�X�^�o�^"
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub
