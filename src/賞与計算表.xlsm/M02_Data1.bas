Attribute VB_Name = "M02_Data1"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private lngR   As Long    '�s����
Private dblR   As Double  '��{�x����
Dim lngKIN(5)  As Long    '���v

Const SQL1 = "SELECT  �Ј��R�[�h, �Ј���, ����, ��{���P, ��{���Q, �Ǘ��E�蓖, �Ƒ��蓖, ����2, ����3, ���喼, �Ј����, �V���Ј� " & _
             "FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪) = '"
Const SQL2 = "') And ((����2) = '"
Const SQL3 = "') And ((�Ј����)='"
Const SQL3Y = "' Or (�Ј����)='"
Const SQL4 = "')) ORDER BY ����3, ���� DESC, �Ј��R�[�h"
Const SQL5 = "') And ((�V���Ј�)<>'Y')) ORDER BY ���� DESC, �Ј��R�[�h"
Const SQL6 = "') And ((�V���Ј�)='Y')) ORDER BY �Ј��R�[�h"
Const SQL7 = "SELECT �������Ə� FROM �O���[�v�Ј��}�X�^�[ WHERE (((���Ə��敪) = '"
Const SQL8 = "')) GROUP BY �������Ə� ORDER BY �������Ə� DESC"

Sub BMN_SET()

Dim strKBN    As String

    Range("AH2:AI22").ClearContents
    
    '�x�X
    strKBN = Range("AE1")
    If strKBN = "TA" Or strKBN = "KA" Then
        strDB = dbT
    Else
        strDB = dbM
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    strSQL = SQL7 & strKBN & SQL8
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 1
    Do Until rsA.EOF
        lngR = lngR + 1
        Cells(lngR, 34) = rsA.Fields(0)
        Select Case rsA.Fields(0)
            Case "OS"
                Cells(lngR, 35) = "���"
            Case "FU"
                Cells(lngR, 35) = "����"
            Case "NG"
                Cells(lngR, 35) = "���É�"
            Case "TK"
                Cells(lngR, 35) = "����"
            Case "SG"
                Cells(lngR, 35) = "��֓�"
            Case "SD"
                Cells(lngR, 35) = "���"
            Case "AK"
                Cells(lngR, 35) = "�k�֓�"
            Case "HB"
                Cells(lngR, 35) = "�{��"
            Case "KA"
                Cells(lngR, 35) = "�֓�"
            Case "TA"
                Cells(lngR, 35) = "���C"
        End Select
        rsA.MoveNext
    Loop
    
    Range("AG1") = 0
    
Exit_DB:
    rsA.Close
    cnA.Close
    Set rsA = Nothing
    Set cnA = Nothing
    
    Call Get_Data
    
End Sub

Sub Get_First()
'=================
'�f�[�^�쐬�{�^��
'=================
Dim strKBN    As String
Dim strBMN    As String
Dim strNXT    As String
Dim strDAT    As String
Dim DateA     As Date
Dim lngC      As Long    '����
Dim lngP      As Long    '�ʒu�L��

    '��ĸر
    With Range("A7:U100")
        .ClearContents
        .Font.Bold = False
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
    With Range("A7:O100")
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    With Range("P7:R100")
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    With Range("B7:I100")
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlHairline
    End With
    With Range("K7:M100")
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlHairline
    End With
    With Range("V7:Z100")
        .ClearContents
    End With
    With Range("E7:F100")
        .NumberFormatLocal = "#,##0"
    End With
    
    '�^�C�g��
    strDAT = Sheets("Main").Range("E2") & "/" & Sheets("Main").Range("G2") & "/10"
    DateA = strDAT
    strDAT = Format(DateA, "ggge") & "�N"
    Range("E4") = strDAT
    If Sheets("Main").Range("G2") = "12" Then
        Range("E4") = Range("E4") & "�~�G"
    ElseIf Sheets("Main").Range("G2") = "7" Then
        Range("E4") = Range("E4") & "�ċG"
    Else
        Range("E4") = Range("E4") & "�Վ�"
    End If
    '�x�X
    strKBN = Range("AE1")
    strDB = dbM
    If Left(strKBN, 1) = "R" Then
        Range("A4") = "�����������Ɗ������ �i" & Range("AF1") & "�j"
    ElseIf strKBN = "KA" Then
        Range("A4") = "�֓��A���R�b�N�H�Ɗ������"
        strDB = dbT
    ElseIf strKBN = "TA" Then
        Range("A4") = "���C�A���R�b�N�H�Ɗ������"
        strDB = dbT
    End If
    
    MsgBox "�V�Ј�������X�V����̂�" & vbCrLf & "�����o�^��ʂœǍ��݁E�o�^��Ƃ����ĉ�����", vbInformation, "�x��"
    
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    lngR = 8
    
    '�c�ƕ��又�� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(7, lngP + 3)
    strSQL = SQL1 & strKBN & SQL2 & "01" & SQL4
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i�c�ƕ���j"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        If strNXT <> rsA.Fields("����3") Then
            lngR = lngR + 1
            Cells(lngR, 1) = "�i" & rsA.Fields("���喼") & "�j"
            lngR = lngR + 1
            strNXT = rsA.Fields("����3")
        End If
        Call ���׏�����F
    Loop
    lngR = lngR + 1
    If strKBN = "TA" Then
        Cells(lngR, 3) = "���Ǘ����升�v"
    Else
        Cells(lngR, 3) = "���c�ƕ��升�v"
    End If
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    rsA.Close
    
    '�H�����又�� ===============================================
    Erase lngKIN
    dblR = Sheets("Main").Cells(8, lngP + 3)
    strSQL = SQL1 & strKBN & SQL2 & "02" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i�H������j"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "���H�����升�v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    End If
    rsA.Close
    
     '���ѕ��又�� ===============================================
    Erase lngKIN
    dblR = Sheets("Main").Cells(9, lngP + 3)
    strSQL = SQL1 & strKBN & SQL2 & "03" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i���ѕ���j"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "�����ѕ��升�v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    End If
    rsA.Close
    
    '�Ǘ����又�� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(10, lngP + 3)
    '����
    Cells(lngR, 1) = "�i�Ǘ�����j"
    Cells(lngR, 1).Font.Bold = True
    lngR = lngR + 2
    Cells(lngR, 1) = "�i��ʎЈ��j"
    Cells(lngR, 6) = "��{(" & dblR & ")"
    '�ް��Ǎ���
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "A" & SQL3Y & "Y" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    rsA.Close
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "B" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        lngR = lngR + 1
        Do Until rsA.EOF
            Call ���׏�����F
        Loop
        lngR = lngR + 1
    End If
    lngR = lngR + 1
    
    Cells(lngR, 3) = "���Ǘ����� �Ј������v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    rsA.Close
    lngR = lngR + 2
    
    '�V���Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(11, lngP + 3)
    '�ް��Ǎ���
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL6
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i�V���Ј��j"
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "���Ǘ����� �V���Ј������v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    End If
    rsA.Close
    
     '�p�[�g�Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(12, lngP + 3)
    '�ް��Ǎ���
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "P" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i�p�[�g�Ј��j"
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "���Ǘ����� �p�[�g�Ј������v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    End If
    rsA.Close
    
     '�����Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(13, lngP + 3)
    '�ް��Ǎ���
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "Z" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i�����Ј��j"
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "���Ǘ����� �����Ј������v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    End If
    rsA.Close
    
    Cells(lngR, 3) = "��" & Range("AC1") & " �����v"
    Cells(lngR, 5) = "=SUM(R7C:R[-1]C)/2"
    Cells(lngR, 6) = "=SUM(R7C:R[-1]C)/2"
    Cells(lngR, 13) = "=SUM(R7C:R[-1]C)/2"
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    
    Range("A1").Select
    
    If Sheets("Main").Range("G2") = "7" And Left(strKBN, 1) <> "R" Then
        MsgBox "�����ɂS���̏��������f���Ă��邩�m�F���鎖"
    End If
    
Exit_DB:
    '�ڑ��̃N���[�Y
    cnA.Close

    '�I�u�W�F�N�g�̔j��
    Set rsA = Nothing
    Set cnA = Nothing

End Sub

Sub ���׏�����F()

        Cells(lngR, 2) = rsA.Fields("�Ј��R�[�h")
        Cells(lngR, 3) = rsA.Fields("�Ј���")
        Cells(lngR, 4) = �����L��(rsA.Fields("����"))
        Cells(lngR, 5) = rsA.Fields("��{���P")
        If IsNull(rsA.Fields("��{���Q")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("��{���Q")
        If IsNull(rsA.Fields("�Ǘ��E�蓖")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("�Ǘ��E�蓖")
        If IsNull(rsA.Fields("�Ƒ��蓖")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("�Ƒ��蓖")
        Cells(lngR, 6) = Application.RoundUp((Cells(lngR, 5) * dblR) / 1000, 0) * 1000
        Cells(lngR, 7) = 0
        Cells(lngR, 8) = "=(RoundUp((RC[-3]*RC[-1])/100,0))*100"
        Cells(lngR, 10) = "/"
        Cells(lngR, 12) = "=IF(RC[1]="""","""",RC[1]-(RC[-6]+RC[-4]))"
        Cells(lngR, 13) = "=IF(RC[-2]=0,RC[-7]+RC[-5],IF(RC[-2]="""","""",RoundUp((((RC[-7]+RC[-5])*RC[-4])/RC[-2])/100,0)*100))"
        Cells(lngR, 14) = "=IF(RC[-1]="""","""",RC[-1]/RC[-9])"
        Cells(lngR, 15) = "=IF(RC[1]=0,"""",(RC[-2]-RC[1])/RC[1])"
        Cells(lngR, 22) = rsA.Fields("����2")
        Cells(lngR, 23) = rsA.Fields("����3")
        Cells(lngR, 24) = rsA.Fields("���喼")
        If rsA.Fields("�Ј����") = "B" Then
            Cells(lngR, 25) = "A"
        ElseIf rsA.Fields("�V���Ј�") = "Y" Then
            Cells(lngR, 1) = "��"
            Cells(lngR, 25) = "Y"
        Else
            Cells(lngR, 25) = rsA.Fields("�Ј����")
        End If
        lngKIN(0) = lngKIN(0) + Cells(lngR, 5)
        lngKIN(1) = lngKIN(1) + Cells(lngR, 6)
        lngKIN(2) = lngKIN(2) + Cells(lngR, 8)
        If Cells(lngR, 13) <> "" Then lngKIN(3) = lngKIN(3) + Cells(lngR, 13)
        If Cells(lngR, 16) <> "" Then lngKIN(4) = lngKIN(4) + Cells(lngR, 16)
        If Cells(lngR, 17) <> "" Then lngKIN(5) = lngKIN(5) + Cells(lngR, 17)
        rsA.MoveNext
        lngR = lngR + 1
        
End Sub

Function �����L��(lngT As Long) As String
    Select Case lngT
        Case 1
            �����L�� = "�T"
        Case 2
            �����L�� = "�U"
        Case 3
            �����L�� = "�V"
        Case 4
            �����L�� = "�W"
        Case 5
            �����L�� = "�X"
        End Select
End Function
