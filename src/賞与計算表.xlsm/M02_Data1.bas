Attribute VB_Name = "M02_Data1"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private Cmd As New ADODB.Command
Private lngR   As Long    '�s����
Private dblR   As Double  '��{�x����
Dim lngKIN(5)  As Long    '���v

Sub BMN_SET()
'===============================
' �����ޯ���x�X�敪�I����Ӽޭ��
' �����ޯ������敪�̑I�������
'===============================
Dim strKBN As String  '���Ə��敪

    Range("AH2:AI22").ClearContents
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    strKBN = Range("AE1")
    strSQL = ""
    strSQL = strSQL & "SELECT OFFICE"
    strSQL = strSQL & "     FROM KYUMTA"
    strSQL = strSQL & "        WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "     GROUP BY OFFICE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
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

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    Call Get_Data
    
End Sub

Sub Get_First()
'=================
'�f�[�^�쐬�{�^��
'=================
Dim strKBN  As String
Dim strBMN  As String
Dim strNXT  As String
Dim strDAT  As String
Dim DateA   As Date
Dim lngC    As Long    '����
Dim lngP    As Long    '�ʒu�L��

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
    If Left(strKBN, 1) = "R" Then
        Range("A4") = "�����������Ɗ������ �i" & Range("AF1") & "�j"
    ElseIf strKBN = "KA" Then
        Range("A4") = "�֓��A���R�b�N�H�Ɗ������"
    ElseIf strKBN = "TA" Then
        Range("A4") = "���C�A���R�b�N�H�Ɗ������"
    End If
    
    MsgBox "�V�Ј�������X�V����̂�" & vbCrLf & "�����o�^��ʂœǍ��݁E�o�^��Ƃ����ĉ�����", vbInformation, "�x��"
    
    strDB = "Initial Catalog=KYUYO;"
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strDB & USER & PSWD 'SQLServer
    cnA.Open
    Set Cmd.ActiveConnection = cnA
    
    '������
    lngR = 8
    Erase lngKIN
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(7, lngP + 3)
    
    '�c�ƕ��又�� ===============================================
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "           And BMN2 = '01'"
    strSQL = strSQL & "    ORDER BY BMN3"
    strSQL = strSQL & "             ,CLASS DESC"
    strSQL = strSQL & "             ,SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    '����
    Cells(lngR, 1) = "�i�c�ƕ���j"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "��{(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        If strNXT <> rsA.Fields("BMN3") Then
            lngR = lngR + 1
            Cells(lngR, 1) = "�i" & Trim(rsA.Fields("BMNNM")) & "�j"
            lngR = lngR + 1
            strNXT = rsA.Fields("BMN3")
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
    Cells(lngR, 13) = lngKIN(2)
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
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "          And BMN2 = '02'"
    strSQL = strSQL & "          And YKBN <> 'Y'"
    strSQL = strSQL & "    ORDER BY CLASS DESC"
    strSQL = strSQL & "              ,SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
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
    Cells(lngR, 13) = lngKIN(2)
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
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "          And BMN2 = '02'"
    strSQL = strSQL & "          And YKBN <> 'Y'"
    strSQL = strSQL & "    ORDER BY CLASS DESC"
    strSQL = strSQL & "              ,SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
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
    Cells(lngR, 13) = lngKIN(2)
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
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "          And BMN2 = '04'"
    strSQL = strSQL & "          And (SKBN ='A' Or SKBN ='B')"
    strSQL = strSQL & "          And YKBN <> 'Y'"
'    strSQL = strSQL & "    ORDER BY CLASS DESC"
'    strSQL = strSQL & "              ,SCODE"
    strSQL = strSQL & "    ORDER BY BMN3"
    strSQL = strSQL & "             ,CLASS DESC"
    strSQL = strSQL & "             ,SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = lngR + 1
    Do Until rsA.EOF
        Call ���׏�����F
    Loop
    rsA.Close
        
    lngR = lngR + 1
    Cells(lngR, 3) = "���Ǘ����� ���v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Cells(lngR, 13) = lngKIN(2)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    lngR = lngR + 2
    
    '�V���Ј����� ===============================================
    Erase lngKIN
    '��{��
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(11, lngP + 3)
    '�ް��Ǎ���
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "           And BMN2 = '04'"
    strSQL = strSQL & "           And (SKBN ='A' Or SKBN ='B')"
    strSQL = strSQL & "           And YKBN = 'Y'"
    strSQL = strSQL & "    ORDER BY SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
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
    Cells(lngR, 3) = "���V���Ј� ���v"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Cells(lngR, 13) = lngKIN(2)
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
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "           And BMN2 = '04'"
    strSQL = strSQL & "           And SKBN = 'P'"
    strSQL = strSQL & "           And YKBN <> 'Y'"
    strSQL = strSQL & "    ORDER BY CLASS DESC"
    strSQL = strSQL & "             ,SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
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
    Cells(lngR, 3) = "���p�[�g�Ј� ���v"
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
    strSQL = ""
    strSQL = strSQL & "SELECT  SCODE"
    strSQL = strSQL & "        ,SNAME"
    strSQL = strSQL & "        ,CLASS"
    strSQL = strSQL & "        ,PAY1"
    strSQL = strSQL & "        ,PAY2"
    strSQL = strSQL & "        ,OPT1"
    strSQL = strSQL & "        ,OPT2"
    strSQL = strSQL & "        ,BMN2"
    strSQL = strSQL & "        ,BMN3"
    strSQL = strSQL & "        ,BMNNM"
    strSQL = strSQL & "        ,SKBN"
    strSQL = strSQL & "        ,YKBN"
    strSQL = strSQL & "    FROM KYUMTA"
    strSQL = strSQL & "         WHERE KBN = '" & strKBN & "'"
    strSQL = strSQL & "           And BMN2 = '04'"
    strSQL = strSQL & "           And SKBN = 'Z'"
    strSQL = strSQL & "           And YKBN <> 'Y'"
    strSQL = strSQL & "    ORDER BY CLASS DESC"
    strSQL = strSQL & "             ,SCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
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
    Cells(lngR, 3) = "�������Ј� ���v"
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

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub

Sub ���׏�����F()

        Cells(lngR, 2) = rsA.Fields("SCODE")
        Cells(lngR, 3) = rsA.Fields("SNAME")
        Cells(lngR, 4) = �����L��(rsA.Fields("CLASS"))
        Cells(lngR, 5) = rsA.Fields("PAY1")
        If IsNull(rsA.Fields("PAY2")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("PAY2")
        If IsNull(rsA.Fields("OPT1")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("OPT1")
        If IsNull(rsA.Fields("OPT2")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("OPT2")
        Cells(lngR, 6) = Application.RoundUp((Cells(lngR, 5) * dblR) / 1000, 0) * 1000
        Cells(lngR, 7) = 0
        Cells(lngR, 8) = "=(RoundUp((RC[-3]*RC[-1])/100,0))*100"
        Cells(lngR, 10) = "/"
        Cells(lngR, 12) = "=IF(RC[1]="""","""",RC[1]-(RC[-6]+RC[-4]))"
        Cells(lngR, 13) = "=IF(RC[-2]=0,RC[-7]+RC[-5],IF(RC[-2]="""","""",RoundUp((((RC[-7]+RC[-5])*RC[-4])/RC[-2])/100,0)*100))"
        Cells(lngR, 14) = "=IF(RC[-1]="""","""",RC[-1]/RC[-9])"
        Cells(lngR, 15) = "=IF(RC[1]=0,"""",(RC[-2]-RC[1])/RC[1])"
        Cells(lngR, 22) = rsA.Fields("BMN2")
        Cells(lngR, 23) = rsA.Fields("BMN3")
        Cells(lngR, 24) = rsA.Fields("BMNNM")
        If rsA.Fields("SKBN") = "B" Then '�Ј��敪(SKBN)��B�̎Ј���A�ɒu��������
            Cells(lngR, 25) = "A"
        ElseIf rsA.Fields("SKBN") = "Y" Then '�V���Ј�����
            Cells(lngR, 1) = "��"
            Cells(lngR, 25) = "Y"
        Else
            Cells(lngR, 25) = rsA.Fields("SKBN")
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
