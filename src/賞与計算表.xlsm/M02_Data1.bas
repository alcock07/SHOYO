Attribute VB_Name = "M02_Data1"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private lngR   As Long    '行ｶｳﾝﾀ
Private dblR   As Double  '基本支給率
Dim lngKIN(5)  As Long    '合計

Const SQL1 = "SELECT  社員コード, 社員名, 等級, 基本給１, 基本給２, 管理職手当, 家族手当, 部門2, 部門3, 部門名, 社員種類, 新入社員 " & _
             "FROM グループ社員マスター WHERE (((事業所区分) = '"
Const SQL2 = "') And ((部門2) = '"
Const SQL3 = "') And ((社員種類)='"
Const SQL3Y = "' Or (社員種類)='"
Const SQL4 = "')) ORDER BY 部門3, 等級 DESC, 社員コード"
Const SQL5 = "') And ((新入社員)<>'Y')) ORDER BY 等級 DESC, 社員コード"
Const SQL6 = "') And ((新入社員)='Y')) ORDER BY 社員コード"
Const SQL7 = "SELECT 所属事業所 FROM グループ社員マスター WHERE (((事業所区分) = '"
Const SQL8 = "')) GROUP BY 所属事業所 ORDER BY 所属事業所 DESC"

Sub BMN_SET()

Dim strKBN    As String

    Range("AH2:AI22").ClearContents
    
    '支店
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
                Cells(lngR, 35) = "大阪"
            Case "FU"
                Cells(lngR, 35) = "福岡"
            Case "NG"
                Cells(lngR, 35) = "名古屋"
            Case "TK"
                Cells(lngR, 35) = "東京"
            Case "SG"
                Cells(lngR, 35) = "南関東"
            Case "SD"
                Cells(lngR, 35) = "仙台"
            Case "AK"
                Cells(lngR, 35) = "北関東"
            Case "HB"
                Cells(lngR, 35) = "本部"
            Case "KA"
                Cells(lngR, 35) = "関東"
            Case "TA"
                Cells(lngR, 35) = "東海"
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
'データ作成ボタン
'=================
Dim strKBN    As String
Dim strBMN    As String
Dim strNXT    As String
Dim strDAT    As String
Dim DateA     As Date
Dim lngC      As Long    '列ｶｳﾝﾀ
Dim lngP      As Long    '位置記憶

    'ｼｰﾄｸﾘｱ
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
    
    'タイトル
    strDAT = Sheets("Main").Range("E2") & "/" & Sheets("Main").Range("G2") & "/10"
    DateA = strDAT
    strDAT = Format(DateA, "ggge") & "年"
    Range("E4") = strDAT
    If Sheets("Main").Range("G2") = "12" Then
        Range("E4") = Range("E4") & "冬季"
    ElseIf Sheets("Main").Range("G2") = "7" Then
        Range("E4") = Range("E4") & "夏季"
    Else
        Range("E4") = Range("E4") & "臨時"
    End If
    '支店
    strKBN = Range("AE1")
    strDB = dbM
    If Left(strKBN, 1) = "R" Then
        Range("A4") = "鳥居金属興業株式会社 （" & Range("AF1") & "）"
    ElseIf strKBN = "KA" Then
        Range("A4") = "関東アルコック工業株式会社"
        strDB = dbT
    ElseIf strKBN = "TA" Then
        Range("A4") = "東海アルコック工業株式会社"
        strDB = dbT
    End If
    
    MsgBox "新社員判定を更新するので" & vbCrLf & "部署登録画面で読込み・登録作業をして下さい", vbInformation, "警告"
    
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    lngR = 8
    
    '営業部門処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(7, lngP + 3)
    strSQL = SQL1 & strKBN & SQL2 & "01" & SQL4
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（営業部門）"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "基本(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        If strNXT <> rsA.Fields("部門3") Then
            lngR = lngR + 1
            Cells(lngR, 1) = "（" & rsA.Fields("部門名") & "）"
            lngR = lngR + 1
            strNXT = rsA.Fields("部門3")
        End If
        Call 明細書込みF
    Loop
    lngR = lngR + 1
    If strKBN = "TA" Then
        Cells(lngR, 3) = "◎管理部門合計"
    Else
        Cells(lngR, 3) = "◎営業部門合計"
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
    
    '工事部門処理 ===============================================
    Erase lngKIN
    dblR = Sheets("Main").Cells(8, lngP + 3)
    strSQL = SQL1 & strKBN & SQL2 & "02" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（工事部門）"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "基本(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call 明細書込みF
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "◎工事部門合計"
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
    
     'ｼｽﾃﾑ部門処理 ===============================================
    Erase lngKIN
    dblR = Sheets("Main").Cells(9, lngP + 3)
    strSQL = SQL1 & strKBN & SQL2 & "03" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（ｼｽﾃﾑ部門）"
    Cells(lngR, 1).Font.Bold = True
    Cells(lngR, 6) = "基本(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call 明細書込みF
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "◎ｼｽﾃﾑ部門合計"
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
    
    '管理部門処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(10, lngP + 3)
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（管理部門）"
    Cells(lngR, 1).Font.Bold = True
    lngR = lngR + 2
    Cells(lngR, 1) = "（一般社員）"
    Cells(lngR, 6) = "基本(" & dblR & ")"
    'ﾃﾞｰﾀ読込み
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "A" & SQL3Y & "Y" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = lngR + 1
    Do Until rsA.EOF
        Call 明細書込みF
    Loop
    rsA.Close
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "B" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        lngR = lngR + 1
        Do Until rsA.EOF
            Call 明細書込みF
        Loop
        lngR = lngR + 1
    End If
    lngR = lngR + 1
    
    Cells(lngR, 3) = "◎管理部門 社員分合計"
    Cells(lngR, 5) = lngKIN(0)
    Cells(lngR, 6) = lngKIN(1)
    Range(Cells(lngR, 1), Cells(lngR, 21)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    rsA.Close
    lngR = lngR + 2
    
    '新入社員処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(11, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL6
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（新入社員）"
    Cells(lngR, 6) = "基本(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call 明細書込みF
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "◎管理部門 新入社員分合計"
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
    
     'パート社員処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(12, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "P" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（パート社員）"
    Cells(lngR, 6) = "基本(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call 明細書込みF
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "◎管理部門 パート社員分合計"
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
    
     '嘱託社員処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(13, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = SQL1 & strKBN & SQL2 & "04" & SQL3 & "Z" & SQL5
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
    rsA.MoveFirst
    'ﾀｲﾄﾙ
    Cells(lngR, 1) = "（嘱託社員）"
    Cells(lngR, 6) = "基本(" & dblR & ")"
    lngR = lngR + 1
    Do Until rsA.EOF
        Call 明細書込みF
    Loop
    lngR = lngR + 1
    Cells(lngR, 3) = "◎管理部門 嘱託社員分合計"
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
    
    Cells(lngR, 3) = "◎" & Range("AC1") & " 総合計"
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
        MsgBox "賃金に４月の昇給が反映しているか確認する事"
    End If
    
Exit_DB:
    '接続のクローズ
    cnA.Close

    'オブジェクトの破棄
    Set rsA = Nothing
    Set cnA = Nothing

End Sub

Sub 明細書込みF()

        Cells(lngR, 2) = rsA.Fields("社員コード")
        Cells(lngR, 3) = rsA.Fields("社員名")
        Cells(lngR, 4) = 等級記号(rsA.Fields("等級"))
        Cells(lngR, 5) = rsA.Fields("基本給１")
        If IsNull(rsA.Fields("基本給２")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("基本給２")
        If IsNull(rsA.Fields("管理職手当")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("管理職手当")
        If IsNull(rsA.Fields("家族手当")) = False Then Cells(lngR, 5) = Cells(lngR, 5) + rsA.Fields("家族手当")
        Cells(lngR, 6) = Application.RoundUp((Cells(lngR, 5) * dblR) / 1000, 0) * 1000
        Cells(lngR, 7) = 0
        Cells(lngR, 8) = "=(RoundUp((RC[-3]*RC[-1])/100,0))*100"
        Cells(lngR, 10) = "/"
        Cells(lngR, 12) = "=IF(RC[1]="""","""",RC[1]-(RC[-6]+RC[-4]))"
        Cells(lngR, 13) = "=IF(RC[-2]=0,RC[-7]+RC[-5],IF(RC[-2]="""","""",RoundUp((((RC[-7]+RC[-5])*RC[-4])/RC[-2])/100,0)*100))"
        Cells(lngR, 14) = "=IF(RC[-1]="""","""",RC[-1]/RC[-9])"
        Cells(lngR, 15) = "=IF(RC[1]=0,"""",(RC[-2]-RC[1])/RC[1])"
        Cells(lngR, 22) = rsA.Fields("部門2")
        Cells(lngR, 23) = rsA.Fields("部門3")
        Cells(lngR, 24) = rsA.Fields("部門名")
        If rsA.Fields("社員種類") = "B" Then
            Cells(lngR, 25) = "A"
        ElseIf rsA.Fields("新入社員") = "Y" Then
            Cells(lngR, 1) = "☆"
            Cells(lngR, 25) = "Y"
        Else
            Cells(lngR, 25) = rsA.Fields("社員種類")
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

Function 等級記号(lngT As Long) As String
    Select Case lngT
        Case 1
            等級記号 = "Ⅰ"
        Case 2
            等級記号 = "Ⅱ"
        Case 3
            等級記号 = "Ⅲ"
        Case 4
            等級記号 = "Ⅳ"
        Case 5
            等級記号 = "Ⅴ"
        End Select
End Function
