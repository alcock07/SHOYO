Attribute VB_Name = "M04_Check"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private strKBN    As String  '�x�X�敪
Private strDAT    As String  '����x���N��
Private strDAZ    As String  '�O��x���N��
Private strDAL    As String  '�O�N�x���N��

'Const SQL1 = "SELECT * FROM �ܗ^ WHERE (((�x���N��) = '"
'Const SQL2 = "') AND ((����1)='"
'Const SQL3 = "') And ((�Ј��R�[�h) = '"
'Const SQL4 = "')) ORDER BY ���� DESC, �Ј��R�[�h"

Sub Proc_Check()

Dim lngR    As Long   '����

    '��ĸر
    Range("A5:E153").ClearContents
    Range("H5:I153").ClearContents
    Range("L5:M153").ClearContents
    
    '�x���N�����
    strDAT = Sheets("Main").Range("E2") & Format(Sheets("Main").Range("G2"), "00")
    If Right(strDAT, 2) = "12" Then
        '�~�G ===============================================
        strDAZ = Sheets("Main").Range("E2") & "07"
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "12"
        Range("D3") = "����(�~�G�j"
        Range("G3") = "�O��(�ċG�j"
        Range("K3") = "�O�N(�~�G�j"
    Else
        '�ċG ===============================================
        strDAZ = CLng(Sheets("Main").Range("E2")) - 1 & "12"
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "07"
        Range("D3") = "����(�ċG�j"
        Range("G3") = "�O��(�~�G�j"
        Range("K3") = "�O�N(�ċG�j"
    End If
    '====================================================
    '���Ə��敪���ƓǍ���
    strKBN = Range("S1")
    If strKBN = "" Then GoTo Exit_DB
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open
'    strSQL = SQL1 & strDAT & SQL2 & strKBN & SQL4
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND ����1 = '" & strKBN & "'"
    strSQL = strSQL & "       ORDER BY ���� DESC"
    strSQL = strSQL & "                ,�Ј��R�[�h"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 5
    Do Until rsA.EOF
        Cells(lngR, 1) = strKBN
        Cells(lngR, 2) = rsA.Fields("�Ј��R�[�h")
        Cells(lngR, 3) = rsA.Fields("�Ј���")
        Cells(lngR, 4) = rsA.Fields("����")
        Cells(lngR, 5) = rsA.Fields("�ܗ^�x���z")
        rsA.MoveNext
        lngR = lngR + 1
    Loop
    
    rsA.Close
    lngR = 5
    Do
        If Cells(lngR, 2) = "" Then Exit Do
'        strSQL = SQL1 & strDAZ & SQL2 & strKBN & SQL3 & Cells(lngR, 2) & SQL4
        strSQL = ""
        strSQL = strSQL & "SELECT *"
        strSQL = strSQL & "       FROM �ܗ^"
        strSQL = strSQL & "            WHERE �x���N�� = '" & strDAZ & "'"
        strSQL = strSQL & "            AND ����1 = '" & strKBN & "'"
        strSQL = strSQL & "            AND �Ј��R�[�h = '" & Cells(lngR, 2) & "'"
        strSQL = strSQL & "       ORDER BY ���� DESC"
        strSQL = strSQL & "              , �Ј��R�[�h"
        rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsA.EOF = False Then
            rsA.MoveFirst
            Cells(lngR, 8) = rsA.Fields("����")
            Cells(lngR, 9) = rsA.Fields("�ܗ^�x���z")
            rsA.MoveNext
        End If
        rsA.Close
        lngR = lngR + 1
    Loop
    
    lngR = 5
    Do
        If Cells(lngR, 2) = "" Then Exit Do
'        strSQL = SQL1 & strDAL & SQL2 & strKBN & SQL3 & Cells(lngR, 2) & SQL4
        strSQL = ""
        strSQL = strSQL & "SELECT *"
        strSQL = strSQL & "       FROM �ܗ^"
        strSQL = strSQL & "            WHERE �x���N�� = '" & strDAL & "'"
        strSQL = strSQL & "            AND ����1 = '" & strKBN & "'"
        strSQL = strSQL & "            AND �Ј��R�[�h = '" & Cells(lngR, 2) & "'"
        strSQL = strSQL & "       ORDER BY ���� DESC"
        strSQL = strSQL & "              , �Ј��R�[�h"
        rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsA.EOF = False Then
            rsA.MoveFirst
            Cells(lngR, 12) = rsA.Fields("����")
            Cells(lngR, 13) = rsA.Fields("�ܗ^�x���z")
            rsA.MoveNext
        End If
        rsA.Close
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
