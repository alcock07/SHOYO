Attribute VB_Name = "M03_Data2"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private strKBN    As String  '�x�X�敪
Private strDAT    As String  '����x���N��
Private strDAL    As String  '�O��x���N��
Private strDAZ    As String  '�O�N�x���N��
Private lngKIN(7) As Long    '���v���z
Private lngRKN(7) As Long    '���v���z
Private lngR      As Long    '�s����
Private dblR      As Double  '��{�x����

'Const SQL2 = "SELECT * FROM �ܗ^ WHERE (((�x���N��) = '"
'Const SQL3 = "') And ((����1) = '"
'Const SQL4 = "') And ((����2) = '"
'Const SQL5 = "') And ((�Ј����) = '"
'Const SQL6 = "') And ((����3) = '"
'Const SQL6S1 = "') And ((����3) > '"
'Const SQL6S2 = "') And ((����3) < '"
'Const SQL8 = "')) ORDER BY ����3, ���� DESC, �Ј��R�[�h"  '�c�ƕ���̂݉ۂ��Ƃɕ��ёւ�
'Const SQL9 = "')) ORDER BY ���� DESC, �Ј��R�[�h"
'Const SQLZ1 = "SELECT �ܗ^�x���z, ���� FROM �ܗ^ WHERE (((�x���N��) = '"
'Const SQLZ2 = "') And ((�Ј��R�[�h) = '"
'Const SQLZ3 = "'))"

Sub Get_Data()
'=================
'�f�[�^�Ǎ��{�^��
'=================
    If Range("AG1") = 0 Then
        Call Proc_Data("S") '�x�X�S���̃��X�g
    Else
        If Range("AD1") < 3 Then
            Call Proc_Data("B") '���傲�Ƃ̃��X�g
        End If
    End If
End Sub

Sub Proc_Data(strSB As String)

Dim strBMN  As String '���喼
Dim strNXT  As String '���唻��p
Dim strMM   As String '��
Dim DateA   As Date   '���t��Ɨp
Dim lngC    As Long   '����
Dim lngP    As Long   '�ʒu�L��
Dim strEg   As String '�c�Ɣ��f
Dim lngErr  As Long   'ٰ�߶���
Dim lngDef  As Long   '��������
Dim lngM    As Long

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
    strDAT = Strings.Format(DateA, "ggge") & "�N"
    Range("E4") = strDAT
    If Sheets("Main").Range("G2") = "12" Then
        Range("E4") = Range("E4") & "�~�G"
    ElseIf Sheets("Main").Range("G2") = "7" Then
        Range("E4") = Range("E4") & "�ċG"
    Else
        Range("E4") = Range("E4") & "�Վ�"
    End If
    '�x���N�����
    strMM = Format(Sheets("Main").Range("G2"), "00")
    strDAT = Sheets("Main").Range("E2") & strMM
    strDAZ = CLng(Sheets("Main").Range("E2")) - 1 & strMM
    If strMM >= "01" And strMM <= "07" Then
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "12"
    Else
        strDAL = Sheets("Main").Range("E2") & "07"
    End If
        
    '�x�X��or���喼�擾
    strKBN = Range("AE1")
    If Range("AG1") > 0 Then strBMN = Range("AH1")
    If Left(strKBN, 1) = "R" Then
        If strSB = "S" Then
            Range("A4") = "�����������Ɗ������ �i" & Range("AF1") & "�j"
        ElseIf strSB = "B" Then
            Range("A4") = "�����������Ɗ������ �i" & Range("AF1") & "-" & Range("AI1") & "�j"
        End If
    ElseIf strKBN = "KA" Then
        Range("A4") = "�֓��A���R�b�N�H�Ɗ������"
    ElseIf strKBN = "TA" Then
        Range("A4") = "���C�A���R�b�N�H�Ɗ������"
    Else
        Range("A4") = ""
    End If
    '�ް��ް������
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open
    lngR = 8
    
    '�c�ƕ��又�� ===============================================
    Erase lngKIN, lngRKN
    '��{��
    lngP = Range("AD1")
    
    If strKBN <> "RH" Then
    
    dblR = Sheets("Main").Cells(7, lngP + 3)
    '�ް��Ǎ�
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND      ����1 = '" & strKBN & "'"
    strSQL = strSQL & "            AND      ����2 = '01'"
    If strSB = "B" Then '���傲�Ƃ̃��X�g
        If strBMN = "OS" Then
           strSQL = strSQL & "  And ����3 > '10'"
           strSQL = strSQL & "  And ����3 < '17'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And ����3 > '24'"
            strSQL = strSQL & "  And ����3 < '27'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And ����3 = '19'"
        ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And ����3 = '22'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And ����3 = '27'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And ����3 = '28'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And ����3 = '29'"
        Else
            GoTo Exit_DB
        End If
    End If
    strSQL = strSQL & "          ORDER BY ����3, ���� DESC, �Ј��R�[�h"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
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
            Call ���׏�����
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "���c�ƕ��升�v"
        Call ���v����
        lngR = lngR + 2
    End If
    rsA.Close
    
    End If
    
    '���ѕ��又�� ===============================================
    Erase lngKIN
    dblR = Sheets("Main").Cells(9, lngP + 3)
    '�ް��Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND      ����1 = '" & strKBN & "'"
    strSQL = strSQL & "            AND      ����2 = '03'"
    If strSB = "B" Then
        strSQL = strSQL & "        And ����3 = '" & strBMN & " '"
    End If
    strSQL = strSQL & "       ORDER BY ���� DESC, �Ј��R�[�h"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        '����
        Cells(lngR, 1) = "�i���ѕ���j"
        Cells(lngR, 1).Font.Bold = True
        Cells(lngR, 6) = "��{(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            Call ���׏�����
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "�����ѕ��升�v"
        Call ���v����
        lngR = lngR + 2
    End If
    rsA.Close
    
    '�c�ƁE�H�����v����
    If strKBN <> "TA" And strKBN <> "KA" Then
        If strEg = "�c��" Then
            If strSB = "S" Then
                Cells(lngR, 3) = "��" & Range("AC1") & " �c�ƥ�H������v"
            ElseIf strSB = "B" Then
                Cells(lngR, 3) = "��" & Range("AF1") & " �c�ƥ�H������v"
            End If
            Cells(lngR, 5) = lngRKN(0)
            Cells(lngR, 6) = lngRKN(1)
            Cells(lngR, 8) = lngRKN(2)
            Cells(lngR, 13) = lngRKN(3)
            Cells(lngR, 14) = "=IF(RC[-1]=0,"""",RC[-1]/RC[-9])"
            Cells(lngR, 15) = "=IF(RC[1]=0,0,RC[-2]/RC[1])"
            Cells(lngR, 16) = lngRKN(4)
            Cells(lngR, 17) = lngRKN(5)
            Cells(lngR, 18) = "=IF(RC[-2]=0,"""",RC[-2]/RC[-1])"
            Cells(lngR, 19) = lngRKN(6)
            Cells(lngR, 20) = lngRKN(7)
            Cells(lngR, 21) = "=IF(RC[-2]=0,"""",RC[-2]/RC[-1])"
            Range(Cells(lngR, 1), Cells(lngR, 21)).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .Weight = xlThick
            End With
            lngR = lngR + 2
        End If
    End If
    
    '�Ǘ����又�� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(10, lngP + 3)
    '�ް��Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND   ����1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   ����2    = '04'"
    strSQL = strSQL & "            AND �@�Ј���� = 'A'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And ����3 > '16'"
            strSQL = strSQL & "  And ����3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And ����3 > '18'"
            strSQL = strSQL & "  And ����3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And ����3 > '22'"
            strSQL = strSQL & "  And ����3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And ����3 > '29'"
            strSQL = strSQL & "  And ����3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And ����3 > '31'"
            strSQL = strSQL & "  And ����3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And ����3 > '33'"
            strSQL = strSQL & "  And ����3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And ����3 > '35'"
            strSQL = strSQL & "  And ����3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY ���� DESC, �Ј��R�[�h"
    Else
        strSQL = strSQL & "  ORDER BY ����3, ���� DESC, �Ј��R�[�h"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        '����
        Cells(lngR, 1) = "�i�Ǘ�����j"
        Cells(lngR, 1).Font.Bold = True
        lngR = lngR + 2
        Cells(lngR, 1) = "�i��ʎЈ��j"
        Cells(lngR, 6) = "��{(" & dblR & ")"
    
        lngR = lngR + 1
        Do Until rsA.EOF
            Call ���׏�����
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "���Ј������v"
        Call ���v����
    End If
    rsA.Close
    lngR = lngR + 2
    
    '�V���Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(11, lngP + 3)
    '�ް��Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND   ����1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   ����2    = '04'"
    strSQL = strSQL & "            AND �@�Ј���� = 'Y'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And ����3 > '16'"
            strSQL = strSQL & "  And ����3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And ����3 > '18'"
            strSQL = strSQL & "  And ����3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And ����3 > '22'"
            strSQL = strSQL & "  And ����3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And ����3 > '29'"
            strSQL = strSQL & "  And ����3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And ����3 > '31'"
            strSQL = strSQL & "  And ����3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And ����3 > '33'"
            strSQL = strSQL & "  And ����3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And ����3 > '35'"
            strSQL = strSQL & "  And ����3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY ���� DESC, �Ј��R�[�h"
    Else
        strSQL = strSQL & "  ORDER BY ����3, ���� DESC, �Ј��R�[�h"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        '����
        Cells(lngR, 1) = "�i�V���Ј��j"
        Cells(lngR, 6) = "��{(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            Call ���׏�����
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "���V���Ј������v"
        Call ���v����
        lngR = lngR + 2
    End If
    rsA.Close

     '�p�[�g�Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(12, lngP + 3)
    '�ް��Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND   ����1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   ����2    = '04'"
    strSQL = strSQL & "            AND �@�Ј���� = 'P'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And ����3 > '16'"
            strSQL = strSQL & "  And ����3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And ����3 > '18'"
            strSQL = strSQL & "  And ����3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And ����3 > '22'"
            strSQL = strSQL & "  And ����3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And ����3 > '29'"
            strSQL = strSQL & "  And ����3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And ����3 > '31'"
            strSQL = strSQL & "  And ����3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And ����3 > '33'"
            strSQL = strSQL & "  And ����3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And ����3 > '35'"
            strSQL = strSQL & "  And ����3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY ���� DESC, �Ј��R�[�h"
    Else
        strSQL = strSQL & "  ORDER BY ����3, ���� DESC, �Ј��R�[�h"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        '����
        Cells(lngR, 1) = "�i�p�[�g�Ј��j"
        Cells(lngR, 6) = "��{(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            If rsA.Fields("���щ�����") = "0.1" Or rsA.Fields("���щ�����") = "-0.1" Then
                lngM = MsgBox("�p�[�g�̐��ї���'0.1' or '-0.1'�ɂȂ��Ă��܂��B�@�m�F���ĉ������B", vbInformation, "���̓`�F�b�N")
            End If
            Call ���׏�����
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "���p�[�g�Ј������v"
        Call ���v����
        lngR = lngR + 2
    End If
    rsA.Close

     '�����Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(13, lngP + 3)
    '�ް��Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND   ����1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   ����2    = '04'"
    strSQL = strSQL & "            AND �@�Ј���� = 'Z'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And ����3 > '16'"
            strSQL = strSQL & "  And ����3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And ����3 > '18'"
            strSQL = strSQL & "  And ����3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And ����3 > '22'"
            strSQL = strSQL & "  And ����3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And ����3 > '29'"
            strSQL = strSQL & "  And ����3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And ����3 > '31'"
            strSQL = strSQL & "  And ����3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And ����3 > '33'"
            strSQL = strSQL & "  And ����3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And ����3 > '35'"
            strSQL = strSQL & "  And ����3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY ���� DESC, �Ј��R�[�h"
    Else
        strSQL = strSQL & "  ORDER BY ����3, ���� DESC, �Ј��R�[�h"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        '����
        Cells(lngR, 1) = "�i�����Ј��j"
        Cells(lngR, 6) = "��{(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            Call ���׏�����
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "�������Ј������v"
        Call ���v����
        lngR = lngR + 2
    End If
    rsA.Close
    
    '�����v����
    If strSB = "S" Then
        Cells(lngR, 3) = "��" & Range("AC1") & " �����v"
    ElseIf strSB = "B" Then
        Cells(lngR, 3) = "��" & Range("AF1") & " �����v"
    End If
    Cells(lngR, 5) = lngRKN(0)
    Cells(lngR, 6) = lngRKN(1)
    Cells(lngR, 8) = lngRKN(2)
    Cells(lngR, 13) = lngRKN(3)
    Cells(lngR, 14) = "=IF(RC[-1]=0,"""",RC[-1]/RC[-9])"
    Cells(lngR, 15) = "=IF(RC[1]=0,0,RC[-2]/RC[1])"
    Cells(lngR, 16) = lngRKN(4)
    Cells(lngR, 17) = lngRKN(5)
    Cells(lngR, 18) = "=IF(RC[-2]=0,"""",RC[-2]/RC[-1])"
    Cells(lngR, 19) = lngRKN(6)
    Cells(lngR, 20) = lngRKN(7)
    Cells(lngR, 21) = "=IF(RC[-2]=0,"""",RC[-2]/RC[-1])"
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    
    lngDef = 0
    If strMM = "07" Then
        For lngErr = 10 To 30
            If (Cells(lngErr, 5) = Cells(lngErr, 17)) And Cells(lngErr, 5) <> "" Then
                lngDef = lngDef + 1
            End If
        Next lngErr
        If lngDef > 4 Then
            MsgBox "�����̏]�ƈ����O��Ɠ��������ł��B" & vbCrLf & "�������`�F�b�N���ĉ������B", vbCritical, "�����`���b�N"
        End If
    End If
    
    Range("A1").Select
    
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

Sub ���׏�����()

Dim rsZ As New ADODB.Recordset

        '�e���ھ��
        Cells(lngR, 2) = rsA.Fields("�Ј��R�[�h")
        Cells(lngR, 3) = rsA.Fields("�Ј���")
        Cells(lngR, 4) = rsA.Fields("����")
        Cells(lngR, 5) = rsA.Fields("����")
        Cells(lngR, 6) = Application.RoundUp((Cells(lngR, 5) * dblR) / 1000, 0) * 1000
        Cells(lngR, 7) = rsA.Fields("���щ�����")
        Cells(lngR, 9) = rsA.Fields("�o�ΌW��1")
        Cells(lngR, 10) = "/"
        Cells(lngR, 11) = rsA.Fields("�o�ΌW��2")
        If rsA.Fields("�Œ�") = "Y" Then
            Cells(lngR, 1) = "     ��"
            Cells(lngR, 8) = 0
            Cells(lngR, 9) = ""
            Cells(lngR, 11) = ""
            Cells(lngR, 13) = rsA.Fields("�ܗ^�x���z")
        ElseIf rsA.Fields("�Œ�") = "S" Then
            Cells(lngR, 1) = "     ��"
            Cells(lngR, 8) = rsA.Fields("���юx���z")
            Cells(lngR, 12) = rsA.Fields("�o�Ό��Z�z")
            Cells(lngR, 13) = rsA.Fields("�ܗ^�x���z")
        Else
            Cells(lngR, 8) = "=RoundUp(((RC[-3]*RC[-1])/100),0)*100"
            Cells(lngR, 13) = "=IF(RC[-2]=0,RC[-7]+RC[-5],IF(RC[-2]="""","""",RoundUp((((RC[-7]+RC[-5])*RC[-4])/RC[-2])/100,0)*100))"
        End If
        Cells(lngR, 12) = "=IF(RC[1]="""","""",RC[1]-(RC[-6]+RC[-4]))"
        Cells(lngR, 14) = "=IF(RC[-1]="""","""",RC[-1]/RC[-9])"
        Cells(lngR, 22) = rsA.Fields("����2")
        Cells(lngR, 23) = rsA.Fields("����3")
        Cells(lngR, 24) = rsA.Fields("���喼")
        Cells(lngR, 25) = rsA.Fields("�Ј����")
        Cells(lngR, 26) = rsA.Fields("�Œ�")
        lngKIN(0) = lngKIN(0) + Cells(lngR, 5)
        lngKIN(1) = lngKIN(1) + Cells(lngR, 6)
        lngKIN(2) = lngKIN(2) + Cells(lngR, 8)
        If Cells(lngR, 13) <> "" Then lngKIN(3) = lngKIN(3) + Cells(lngR, 13)
        '�O���ް��擾
        strSQL = ""
        strSQL = strSQL & "SELECT �ܗ^�x���z"
        strSQL = strSQL & "      ,����"
        strSQL = strSQL & "       FROM �ܗ^"
        strSQL = strSQL & "            WHERE �x���N�� = '" & strDAL & "'"
        strSQL = strSQL & "            AND   �Ј��R�[�h = '" & rsA.Fields("�Ј��R�[�h") & "'"
        rsZ.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsZ.EOF = False Then
            Cells(lngR, 15) = "=IF(RC[1]=0,0,RC[-2]/RC[1])"
            Cells(lngR, 16) = rsZ.Fields("�ܗ^�x���z")
            Cells(lngR, 17) = rsZ.Fields("����")
            Cells(lngR, 18) = "=IF(RC[-2]=0,0,RC[-2]/RC[-1])"
            lngKIN(4) = lngKIN(4) + Cells(lngR, 16)
            lngKIN(5) = lngKIN(5) + Cells(lngR, 17)
        End If
        rsZ.Close
        '�O�N�ް��擾
        strSQL = ""
        strSQL = strSQL & "SELECT �ܗ^�x���z"
        strSQL = strSQL & "      ,����"
        strSQL = strSQL & "       FROM �ܗ^"
        strSQL = strSQL & "            WHERE �x���N�� = '" & strDAZ & "'"
        strSQL = strSQL & "            AND   �Ј��R�[�h = '" & rsA.Fields("�Ј��R�[�h") & "'"
        rsZ.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsZ.EOF = False Then
            Cells(lngR, 19) = rsZ.Fields("�ܗ^�x���z")
            Cells(lngR, 20) = rsZ.Fields("����")
            Cells(lngR, 21) = "=IF(RC[-2]=0,0,RC[-2]/RC[-1])"
            lngKIN(6) = lngKIN(6) + Cells(lngR, 19)
            lngKIN(7) = lngKIN(7) + Cells(lngR, 20)
        End If
        rsZ.Close
        rsA.MoveNext
        If Range("AG1") = 0 Then
            lngR = lngR + 1
        Else
            lngR = lngR + 2
        End If
End Sub

Sub ���v����()
Dim lngI As Long

    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Cells(lngR, 8) = lngKIN(2)
    Cells(lngR, 13) = lngKIN(3)
    Cells(lngR, 14) = "=IF(RC[-1]="""","""",RC[-1]/RC[-9])"
    Cells(lngR, 15) = "=IF(RC[1]=0,0,RC[-2]/RC[1])"
    Cells(lngR, 16) = lngKIN(4)
    Cells(lngR, 17) = lngKIN(5)
    Cells(lngR, 18) = "=IF(RC[-2]=0,"""",RC[-2]/RC[-1])"
    Cells(lngR, 19) = lngKIN(6)
    Cells(lngR, 20) = lngKIN(7)
    Cells(lngR, 21) = "=IF(RC[-2]=0,"""",RC[-2]/RC[-1])"

    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    For lngI = 0 To 7
        lngRKN(lngI) = lngRKN(lngI) + lngKIN(lngI)
    Next lngI
    
End Sub

Sub Up_Data()
'=================
'�f�[�^�o�^�{�^��
'=================
'Const SQLD1 = "DELETE FROM �ܗ^ WHERE (((�x���N��)='"
'Const SQLD2 = "') AND ((����1)='"
'Const SQLD3 = "'))"
Const SQL1 = "SELECT * FROM �ܗ^"

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strDAT   As String
Dim strKBN   As String
Dim DateA    As Date
Dim lngR     As Long    '�s����
Dim lngC     As Long    '����


    For lngR = 8 To 100
        If Cells(lngR, 3) = "�� �����v" Then Exit For
        If Cells(lngR, 11) <> "" Then
            If Cells(lngR, 9) > Cells(lngR, 11) Then
                MsgBox "�o�ΌW�����m�F���ĉ������I " & lngR & "�s��", vbCritical
                MsgBox "�o�^���s�I(T-T)", vbExclamation, "�o�^"
                GoTo Exit_DB
            End If
        End If
    Next lngR
    If Range("AG1") <> 0 Then
        MsgBox "�o�^�͕��傲�Ƃɂ͏o���܂���Im(__)m", vbCritical, "�o�^�G���["
        MsgBox "�o�^���s�I(T-T)", vbExclamation, "�o�^"
        GoTo Exit_DB
    End If
    
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open

    '�f�[�^�폜����
    strDAT = Sheets("Main").Range("E2") & Format(Sheets("Main").Range("G2"), "00")
    strKBN = Range("AE1")
'    strSQL = SQLD1 & strDAT & SQLD2 & strKBN & SQLD3
    strSQL = ""
    strSQL = strSQL & "DELETE"
    strSQL = strSQL & "       FROM �ܗ^"
    strSQL = strSQL & "            WHERE �x���N�� = '" & strDAT & "'"
    strSQL = strSQL & "            AND   ����1    = '" & strKBN & "'"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    
    '�f�[�^�o�^����
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM �ܗ^"
    rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
    For lngR = 8 To 100
        If Cells(lngR, 2) <> "" Then
            rsA.AddNew
            rsA.Fields("�x���N��") = strDAT
            rsA.Fields("�Ј��R�[�h") = Cells(lngR, 2)
            rsA.Fields("�Ј���") = Cells(lngR, 3)
            rsA.Fields("����") = Cells(lngR, 4)
            rsA.Fields("����") = Cells(lngR, 5)
            rsA.Fields("��{�x���z") = Cells(lngR, 6)
            rsA.Fields("���щ�����") = Cells(lngR, 7)
            rsA.Fields("���юx���z") = Cells(lngR, 8)
            rsA.Fields("�o�ΌW��1") = Cells(lngR, 9)
            rsA.Fields("�o�ΌW��2") = Cells(lngR, 11)
            If IsError(Cells(lngR, 12)) = False Then
                rsA.Fields("�o�Ό��Z�z") = Cells(lngR, 12)
            End If
            If IsError(Cells(lngR, 13)) = False Then
                rsA.Fields("�ܗ^�x���z") = Cells(lngR, 13)
            End If
            rsA.Fields("����1") = strKBN
            rsA.Fields("����2") = Cells(lngR, 22)
            rsA.Fields("����3") = Cells(lngR, 23)
            rsA.Fields("���喼") = Cells(lngR, 24)
            rsA.Fields("�Ј����") = Cells(lngR, 25)
            rsA.Fields("�Œ�") = Cells(lngR, 26)
            rsA.Update
        End If
    Next lngR
    
    MsgBox "�o�^���܂���(��_��;)", vbExclamation, "�o�^"
        
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

Sub Proc_Prn(boolPDF As Boolean)

    Dim strCPN As String
    Dim strDP   As String
    Dim strJP   As String
    Dim lngM   As Long
    
    strDP = Application.ActivePrinter
    strCPN = CP_NAME
    
    If boolPDF = True Then
        If strCPN = "HB14" Then
            strJP = "Adobe PDF on Ne07:"
        End If
    Else
        If strCPN = "HB14" Then
            strJP = "IR-ADVC5850F on Ne02:"
        End If
    End If
    
    For lngR = 8 To 100
        If InStr(1, Cells(lngR, 3), "�����v") <> 0 Then
            Exit For
        End If
    Next lngR
    
    lngM = MsgBox("�������O�ɕK���o�^�E�Ǎ����鎖�I" & vbCrLf & "������܂����H", vbYesNo, "�v�Z�\���")
    If lngM = vbYes Then
                 
        strDP = Application.ActivePrinter
        Application.ActivePrinter = strJP
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.ResetAllPageBreaks

        With ActiveSheet.PageSetup
                .PrintArea = "$A$4:$U$" & CStr(lngR)
                If boolPDF = False Then .PaperSize = xlPaperB4
                .Zoom = 89
        End With
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:=strJP, Collate:=True
        ActiveSheet.DisplayPageBreaks = False
        ActiveWindow.View = xlNormalView
        Application.ActivePrinter = strDP
        Range("A1").Select
        
    End If
    
End Sub


Sub Prn_Canon()
    Call Proc_Prn(False)
End Sub

Sub Prn_PDF()
    Call Proc_Prn(True)
End Sub
