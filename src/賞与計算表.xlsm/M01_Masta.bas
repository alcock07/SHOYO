Attribute VB_Name = "M01_Masta"
Option Explicit

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

Sub Get_Masta()

Const SQL1 = "SELECT ���Ə��敪, �Ј��R�[�h, �Ј���, �Ј����, ����, ��{���P, ��{���Q, �Ǘ��E�蓖, �Ƒ��蓖, ����1, ����2, ����3, ���喼, ���ДN���� FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪)='"
Const SQL2 = "'))"

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strKBN As String
Dim lngR   As Long
Dim lngC   As Long
Dim DateA  As Date
Dim DateB  As Date
Dim strYY  As String
Dim lngMM  As Long

    Range("A4:J152").ClearContents
    Range("L4:L52").ClearContents
    Range("N4:N52").ClearContents
    
    strKBN = Range("Q2")
    If strKBN = "" Then GoTo Exit_DB
    If strKBN = "TA" Or strKBN = "KA" Then
        strDB = dbT
    Else
        strDB = dbM
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    
    '���Ə��敪���ƓǍ���
    strSQL = SQL1 & strKBN & SQL2
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 4
    Do Until rsA.EOF
        '�e���ھ��
        For lngC = 0 To 8
            Cells(lngR, lngC + 1) = rsA.Fields(lngC)
        Next lngC
        '����敪���
        If IsNull(rsA.Fields("����2")) = False Then Cells(lngR, 10) = rsA.Fields("����2")
        If IsNull(rsA.Fields("����3")) = False Then Cells(lngR, 12) = rsA.Fields("����3")
        '���N����
        If rsA.Fields("���ДN����") <> "" Then
            DateA = rsA.Fields("���ДN����")
        End If
        '�V���Ј����菈��
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
    rsA.Close
    cnA.Close

    Set rsA = Nothing
    Set cnA = Nothing

End Sub

Sub Up_Masta()

Const SQL1 = "SELECT ����1, ����2, ����3, ���喼, �V���Ј� FROM �O���[�v�Ј��}�X�^�[ WHERE (((�Ј��R�[�h)='"
Const SQL2 = "'))"


Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strKBN As String
Dim strCD  As String
Dim strKB1 As String
Dim strKB2 As String
Dim strKB3 As String
Dim lngR   As Long
Dim lngC   As Long
    
    strKBN = Range("Q2")
    If strKBN = "" Then GoTo Exit_DB
    If strKBN = "TA" Or strKBN = "KA" Then
        strDB = dbT
    Else
        strDB = dbM
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open

    lngR = 4
    Do
        strCD = Cells(lngR, 2) '�Ј�����
        If strCD = "" Then Exit Do
        strKB1 = Range("P2")
        strKB2 = Cells(lngR, 10)
        strKB3 = Cells(lngR, 12)
        If strCD <> "" Then
            'Ͻ��ďo
            strSQL = SQL1 & strCD & SQL2
            rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If rsA.EOF = False Then
                rsA.MoveFirst
                rsA.Fields(0) = strKB1
                rsA.Fields(1) = strKB2
                rsA.Fields(2) = strKB3
                rsA.Fields(3) = Cells(lngR, 13)
                If Cells(lngR, 14) = "��" Then
                    rsA.Fields(4) = "Y"
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

    '�ڑ��̃N���[�Y
    cnA.Close

    '�I�u�W�F�N�g�̔j��
    Set rsA = Nothing
    Set cnA = Nothing

End Sub
