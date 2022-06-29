Attribute VB_Name = "M03_Data2"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private strKBN    As String  '支店区分
Private strDAT    As String  '今回支給年月
Private strDAL    As String  '前回支給年月
Private strDAZ    As String  '前年支給年月
Private lngKIN(7) As Long    '小計金額
Private lngRKN(7) As Long    '合計金額
Private lngR      As Long    '行ｶｳﾝﾀ
Private dblR      As Double  '基本支給率

'Const SQL2 = "SELECT * FROM 賞与 WHERE (((支給年月) = '"
'Const SQL3 = "') And ((部門1) = '"
'Const SQL4 = "') And ((部門2) = '"
'Const SQL5 = "') And ((社員種類) = '"
'Const SQL6 = "') And ((部門3) = '"
'Const SQL6S1 = "') And ((部門3) > '"
'Const SQL6S2 = "') And ((部門3) < '"
'Const SQL8 = "')) ORDER BY 部門3, 等級 DESC, 社員コード"  '営業部門のみ課ごとに並び替え
'Const SQL9 = "')) ORDER BY 等級 DESC, 社員コード"
'Const SQLZ1 = "SELECT 賞与支給額, 賃金 FROM 賞与 WHERE (((支給年月) = '"
'Const SQLZ2 = "') And ((社員コード) = '"
'Const SQLZ3 = "'))"

Sub Get_Data()
'=================
'データ読込ボタン
'=================
    If Range("AG1") = 0 Then
        Call Proc_Data("S") '支店全部のリスト
    Else
        If Range("AD1") < 3 Then
            Call Proc_Data("B") '部門ごとのリスト
        End If
    End If
End Sub

Sub Proc_Data(strSB As String)

Dim strBMN  As String '部門名
Dim strNXT  As String '部門判定用
Dim strMM   As String '月
Dim DateA   As Date   '日付作業用
Dim lngC    As Long   '列ｶｳﾝﾀ
Dim lngP    As Long   '位置記憶
Dim strEg   As String '営業判断
Dim lngErr  As Long   'ﾙｰﾌﾟｶｳﾝﾀ
Dim lngDef  As Long   'ﾁｪｯｸｶｳﾝﾀ
Dim lngM    As Long

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
    strDAT = Strings.Format(DateA, "ggge") & "年"
    Range("E4") = strDAT
    If Sheets("Main").Range("G2") = "12" Then
        Range("E4") = Range("E4") & "冬季"
    ElseIf Sheets("Main").Range("G2") = "7" Then
        Range("E4") = Range("E4") & "夏季"
    Else
        Range("E4") = Range("E4") & "臨時"
    End If
    '支給年月ｾｯﾄ
    strMM = Format(Sheets("Main").Range("G2"), "00")
    strDAT = Sheets("Main").Range("E2") & strMM
    strDAZ = CLng(Sheets("Main").Range("E2")) - 1 & strMM
    If strMM >= "01" And strMM <= "07" Then
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "12"
    Else
        strDAL = Sheets("Main").Range("E2") & "07"
    End If
        
    '支店名or部門名取得
    strKBN = Range("AE1")
    If Range("AG1") > 0 Then strBMN = Range("AH1")
    If Left(strKBN, 1) = "R" Then
        If strSB = "S" Then
            Range("A4") = "鳥居金属興業株式会社 （" & Range("AF1") & "）"
        ElseIf strSB = "B" Then
            Range("A4") = "鳥居金属興業株式会社 （" & Range("AF1") & "-" & Range("AI1") & "）"
        End If
    ElseIf strKBN = "KA" Then
        Range("A4") = "関東アルコック工業株式会社"
    ElseIf strKBN = "TA" Then
        Range("A4") = "東海アルコック工業株式会社"
    Else
        Range("A4") = ""
    End If
    'ﾃﾞｰﾀﾍﾞｰｽｵｰﾌﾟﾝ
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open
    lngR = 8
    
    '営業部門処理 ===============================================
    Erase lngKIN, lngRKN
    '基本率
    lngP = Range("AD1")
    
    If strKBN <> "RH" Then
    
    dblR = Sheets("Main").Cells(7, lngP + 3)
    'ﾃﾞｰﾀ読込
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND      部門1 = '" & strKBN & "'"
    strSQL = strSQL & "            AND      部門2 = '01'"
    If strSB = "B" Then '部門ごとのリスト
        If strBMN = "OS" Then
           strSQL = strSQL & "  And 部門3 > '10'"
           strSQL = strSQL & "  And 部門3 < '17'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And 部門3 > '24'"
            strSQL = strSQL & "  And 部門3 < '27'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And 部門3 = '19'"
        ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And 部門3 = '22'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And 部門3 = '27'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And 部門3 = '28'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And 部門3 = '29'"
        Else
            GoTo Exit_DB
        End If
    End If
    strSQL = strSQL & "          ORDER BY 部門3, 等級 DESC, 社員コード"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
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
            Call 明細書込み
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "◎営業部門合計"
        Call 小計処理
        lngR = lngR + 2
    End If
    rsA.Close
    
    End If
    
    'ｼｽﾃﾑ部門処理 ===============================================
    Erase lngKIN
    dblR = Sheets("Main").Cells(9, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND      部門1 = '" & strKBN & "'"
    strSQL = strSQL & "            AND      部門2 = '03'"
    If strSB = "B" Then
        strSQL = strSQL & "        And 部門3 = '" & strBMN & " '"
    End If
    strSQL = strSQL & "       ORDER BY 等級 DESC, 社員コード"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        'ﾀｲﾄﾙ
        Cells(lngR, 1) = "（ｼｽﾃﾑ部門）"
        Cells(lngR, 1).Font.Bold = True
        Cells(lngR, 6) = "基本(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            Call 明細書込み
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "◎ｼｽﾃﾑ部門合計"
        Call 小計処理
        lngR = lngR + 2
    End If
    rsA.Close
    
    '営業・工事合計処理
    If strKBN <> "TA" And strKBN <> "KA" Then
        If strEg = "営業" Then
            If strSB = "S" Then
                Cells(lngR, 3) = "◎" & Range("AC1") & " 営業･工事部門計"
            ElseIf strSB = "B" Then
                Cells(lngR, 3) = "◎" & Range("AF1") & " 営業･工事部門計"
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
    
    '管理部門処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(10, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND   部門1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   部門2    = '04'"
    strSQL = strSQL & "            AND 　社員種類 = 'A'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And 部門3 > '16'"
            strSQL = strSQL & "  And 部門3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And 部門3 > '18'"
            strSQL = strSQL & "  And 部門3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And 部門3 > '22'"
            strSQL = strSQL & "  And 部門3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And 部門3 > '29'"
            strSQL = strSQL & "  And 部門3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And 部門3 > '31'"
            strSQL = strSQL & "  And 部門3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And 部門3 > '33'"
            strSQL = strSQL & "  And 部門3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And 部門3 > '35'"
            strSQL = strSQL & "  And 部門3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY 等級 DESC, 社員コード"
    Else
        strSQL = strSQL & "  ORDER BY 部門3, 等級 DESC, 社員コード"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        'ﾀｲﾄﾙ
        Cells(lngR, 1) = "（管理部門）"
        Cells(lngR, 1).Font.Bold = True
        lngR = lngR + 2
        Cells(lngR, 1) = "（一般社員）"
        Cells(lngR, 6) = "基本(" & dblR & ")"
    
        lngR = lngR + 1
        Do Until rsA.EOF
            Call 明細書込み
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "◎社員分合計"
        Call 小計処理
    End If
    rsA.Close
    lngR = lngR + 2
    
    '新入社員処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(11, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND   部門1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   部門2    = '04'"
    strSQL = strSQL & "            AND 　社員種類 = 'Y'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And 部門3 > '16'"
            strSQL = strSQL & "  And 部門3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And 部門3 > '18'"
            strSQL = strSQL & "  And 部門3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And 部門3 > '22'"
            strSQL = strSQL & "  And 部門3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And 部門3 > '29'"
            strSQL = strSQL & "  And 部門3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And 部門3 > '31'"
            strSQL = strSQL & "  And 部門3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And 部門3 > '33'"
            strSQL = strSQL & "  And 部門3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And 部門3 > '35'"
            strSQL = strSQL & "  And 部門3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY 等級 DESC, 社員コード"
    Else
        strSQL = strSQL & "  ORDER BY 部門3, 等級 DESC, 社員コード"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        'ﾀｲﾄﾙ
        Cells(lngR, 1) = "（新入社員）"
        Cells(lngR, 6) = "基本(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            Call 明細書込み
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "◎新入社員分合計"
        Call 小計処理
        lngR = lngR + 2
    End If
    rsA.Close

     'パート社員処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(12, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND   部門1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   部門2    = '04'"
    strSQL = strSQL & "            AND 　社員種類 = 'P'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And 部門3 > '16'"
            strSQL = strSQL & "  And 部門3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And 部門3 > '18'"
            strSQL = strSQL & "  And 部門3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And 部門3 > '22'"
            strSQL = strSQL & "  And 部門3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And 部門3 > '29'"
            strSQL = strSQL & "  And 部門3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And 部門3 > '31'"
            strSQL = strSQL & "  And 部門3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And 部門3 > '33'"
            strSQL = strSQL & "  And 部門3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And 部門3 > '35'"
            strSQL = strSQL & "  And 部門3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY 等級 DESC, 社員コード"
    Else
        strSQL = strSQL & "  ORDER BY 部門3, 等級 DESC, 社員コード"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        'ﾀｲﾄﾙ
        Cells(lngR, 1) = "（パート社員）"
        Cells(lngR, 6) = "基本(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            If rsA.Fields("成績加減率") = "0.1" Or rsA.Fields("成績加減率") = "-0.1" Then
                lngM = MsgBox("パートの成績率が'0.1' or '-0.1'になっています。　確認して下さい。", vbInformation, "入力チェック")
            End If
            Call 明細書込み
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "◎パート社員分合計"
        Call 小計処理
        lngR = lngR + 2
    End If
    rsA.Close

     '嘱託社員処理 ===============================================
    Erase lngKIN
    '基本率
    lngP = Range("AD1")
    dblR = Sheets("Main").Cells(13, lngP + 3)
    'ﾃﾞｰﾀ読込み
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND   部門1    = '" & strKBN & "'"
    strSQL = strSQL & "            AND   部門2    = '04'"
    strSQL = strSQL & "            AND 　社員種類 = 'Z'"
    If strSB = "B" Then
        If strBMN = "OS" Then
            strSQL = strSQL & "  And 部門3 > '16'"
            strSQL = strSQL & "  And 部門3 < '19'"
        ElseIf strBMN = "FU" Then
            strSQL = strSQL & "  And 部門3 > '18'"
            strSQL = strSQL & "  And 部門3 < '22'"
         ElseIf strBMN = "NG" Then
            strSQL = strSQL & "  And 部門3 > '22'"
            strSQL = strSQL & "  And 部門3 < '25'"
        ElseIf strBMN = "TK" Then
            strSQL = strSQL & "  And 部門3 > '29'"
            strSQL = strSQL & "  And 部門3 < '32'"
        ElseIf strBMN = "SG" Then
            strSQL = strSQL & "  And 部門3 > '31'"
            strSQL = strSQL & "  And 部門3 < '34'"
        ElseIf strBMN = "SD" Then
            strSQL = strSQL & "  And 部門3 > '33'"
            strSQL = strSQL & "  And 部門3 < '36'"
        ElseIf strBMN = "AK" Then
            strSQL = strSQL & "  And 部門3 > '35'"
            strSQL = strSQL & "  And 部門3 < '38'"
        End If
    End If
    If strSB = "S" Then
        strSQL = strSQL & "  ORDER BY 等級 DESC, 社員コード"
    Else
        strSQL = strSQL & "  ORDER BY 部門3, 等級 DESC, 社員コード"
    End If
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then
        rsA.MoveFirst
        'ﾀｲﾄﾙ
        Cells(lngR, 1) = "（嘱託社員）"
        Cells(lngR, 6) = "基本(" & dblR & ")"
        lngR = lngR + 1
        Do Until rsA.EOF
            Call 明細書込み
        Loop
        lngR = lngR + 1
        Cells(lngR, 3) = "◎嘱託社員分合計"
        Call 小計処理
        lngR = lngR + 2
    End If
    rsA.Close
    
    '総合計処理
    If strSB = "S" Then
        Cells(lngR, 3) = "◎" & Range("AC1") & " 総合計"
    ElseIf strSB = "B" Then
        Cells(lngR, 3) = "◎" & Range("AF1") & " 総合計"
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
            MsgBox "複数の従業員が前回と同じ賃金です。" & vbCrLf & "賃金をチェックして下さい。", vbCritical, "賃金チャック"
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

Sub 明細書込み()

Dim rsZ As New ADODB.Recordset

        '各項目ｾｯﾄ
        Cells(lngR, 2) = rsA.Fields("社員コード")
        Cells(lngR, 3) = rsA.Fields("社員名")
        Cells(lngR, 4) = rsA.Fields("等級")
        Cells(lngR, 5) = rsA.Fields("賃金")
        Cells(lngR, 6) = Application.RoundUp((Cells(lngR, 5) * dblR) / 1000, 0) * 1000
        Cells(lngR, 7) = rsA.Fields("成績加減率")
        Cells(lngR, 9) = rsA.Fields("出勤係数1")
        Cells(lngR, 10) = "/"
        Cells(lngR, 11) = rsA.Fields("出勤係数2")
        If rsA.Fields("固定") = "Y" Then
            Cells(lngR, 1) = "     ☆"
            Cells(lngR, 8) = 0
            Cells(lngR, 9) = ""
            Cells(lngR, 11) = ""
            Cells(lngR, 13) = rsA.Fields("賞与支給額")
        ElseIf rsA.Fields("固定") = "S" Then
            Cells(lngR, 1) = "     ☆"
            Cells(lngR, 8) = rsA.Fields("成績支給額")
            Cells(lngR, 12) = rsA.Fields("出勤減算額")
            Cells(lngR, 13) = rsA.Fields("賞与支給額")
        Else
            Cells(lngR, 8) = "=RoundUp(((RC[-3]*RC[-1])/100),0)*100"
            Cells(lngR, 13) = "=IF(RC[-2]=0,RC[-7]+RC[-5],IF(RC[-2]="""","""",RoundUp((((RC[-7]+RC[-5])*RC[-4])/RC[-2])/100,0)*100))"
        End If
        Cells(lngR, 12) = "=IF(RC[1]="""","""",RC[1]-(RC[-6]+RC[-4]))"
        Cells(lngR, 14) = "=IF(RC[-1]="""","""",RC[-1]/RC[-9])"
        Cells(lngR, 22) = rsA.Fields("部門2")
        Cells(lngR, 23) = rsA.Fields("部門3")
        Cells(lngR, 24) = rsA.Fields("部門名")
        Cells(lngR, 25) = rsA.Fields("社員種類")
        Cells(lngR, 26) = rsA.Fields("固定")
        lngKIN(0) = lngKIN(0) + Cells(lngR, 5)
        lngKIN(1) = lngKIN(1) + Cells(lngR, 6)
        lngKIN(2) = lngKIN(2) + Cells(lngR, 8)
        If Cells(lngR, 13) <> "" Then lngKIN(3) = lngKIN(3) + Cells(lngR, 13)
        '前回ﾃﾞｰﾀ取得
        strSQL = ""
        strSQL = strSQL & "SELECT 賞与支給額"
        strSQL = strSQL & "      ,賃金"
        strSQL = strSQL & "       FROM 賞与"
        strSQL = strSQL & "            WHERE 支給年月 = '" & strDAL & "'"
        strSQL = strSQL & "            AND   社員コード = '" & rsA.Fields("社員コード") & "'"
        rsZ.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsZ.EOF = False Then
            Cells(lngR, 15) = "=IF(RC[1]=0,0,RC[-2]/RC[1])"
            Cells(lngR, 16) = rsZ.Fields("賞与支給額")
            Cells(lngR, 17) = rsZ.Fields("賃金")
            Cells(lngR, 18) = "=IF(RC[-2]=0,0,RC[-2]/RC[-1])"
            lngKIN(4) = lngKIN(4) + Cells(lngR, 16)
            lngKIN(5) = lngKIN(5) + Cells(lngR, 17)
        End If
        rsZ.Close
        '前年ﾃﾞｰﾀ取得
        strSQL = ""
        strSQL = strSQL & "SELECT 賞与支給額"
        strSQL = strSQL & "      ,賃金"
        strSQL = strSQL & "       FROM 賞与"
        strSQL = strSQL & "            WHERE 支給年月 = '" & strDAZ & "'"
        strSQL = strSQL & "            AND   社員コード = '" & rsA.Fields("社員コード") & "'"
        rsZ.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsZ.EOF = False Then
            Cells(lngR, 19) = rsZ.Fields("賞与支給額")
            Cells(lngR, 20) = rsZ.Fields("賃金")
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

Sub 小計処理()
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
'データ登録ボタン
'=================
'Const SQLD1 = "DELETE FROM 賞与 WHERE (((支給年月)='"
'Const SQLD2 = "') AND ((部門1)='"
'Const SQLD3 = "'))"
Const SQL1 = "SELECT * FROM 賞与"

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strDAT   As String
Dim strKBN   As String
Dim DateA    As Date
Dim lngR     As Long    '行ｶｳﾝﾀ
Dim lngC     As Long    '列ｶｳﾝﾀ


    For lngR = 8 To 100
        If Cells(lngR, 3) = "◎ 総合計" Then Exit For
        If Cells(lngR, 11) <> "" Then
            If Cells(lngR, 9) > Cells(lngR, 11) Then
                MsgBox "出勤係数を確認して下さい！ " & lngR & "行目", vbCritical
                MsgBox "登録失敗！(T-T)", vbExclamation, "登録"
                GoTo Exit_DB
            End If
        End If
    Next lngR
    If Range("AG1") <> 0 Then
        MsgBox "登録は部門ごとには出来ません！m(__)m", vbCritical, "登録エラー"
        MsgBox "登録失敗！(T-T)", vbExclamation, "登録"
        GoTo Exit_DB
    End If
    
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open

    'データ削除処理
    strDAT = Sheets("Main").Range("E2") & Format(Sheets("Main").Range("G2"), "00")
    strKBN = Range("AE1")
'    strSQL = SQLD1 & strDAT & SQLD2 & strKBN & SQLD3
    strSQL = ""
    strSQL = strSQL & "DELETE"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND   部門1    = '" & strKBN & "'"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    
    'データ登録処理
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
    For lngR = 8 To 100
        If Cells(lngR, 2) <> "" Then
            rsA.AddNew
            rsA.Fields("支給年月") = strDAT
            rsA.Fields("社員コード") = Cells(lngR, 2)
            rsA.Fields("社員名") = Cells(lngR, 3)
            rsA.Fields("等級") = Cells(lngR, 4)
            rsA.Fields("賃金") = Cells(lngR, 5)
            rsA.Fields("基本支給額") = Cells(lngR, 6)
            rsA.Fields("成績加減率") = Cells(lngR, 7)
            rsA.Fields("成績支給額") = Cells(lngR, 8)
            rsA.Fields("出勤係数1") = Cells(lngR, 9)
            rsA.Fields("出勤係数2") = Cells(lngR, 11)
            If IsError(Cells(lngR, 12)) = False Then
                rsA.Fields("出勤減算額") = Cells(lngR, 12)
            End If
            If IsError(Cells(lngR, 13)) = False Then
                rsA.Fields("賞与支給額") = Cells(lngR, 13)
            End If
            rsA.Fields("部門1") = strKBN
            rsA.Fields("部門2") = Cells(lngR, 22)
            rsA.Fields("部門3") = Cells(lngR, 23)
            rsA.Fields("部門名") = Cells(lngR, 24)
            rsA.Fields("社員種類") = Cells(lngR, 25)
            rsA.Fields("固定") = Cells(lngR, 26)
            rsA.Update
        End If
    Next lngR
    
    MsgBox "登録しました(＠_＠;)", vbExclamation, "登録"
        
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
        If InStr(1, Cells(lngR, 3), "総合計") <> 0 Then
            Exit For
        End If
    Next lngR
    
    lngM = MsgBox("印刷する前に必ず登録・読込する事！" & vbCrLf & "印刷しますか？", vbYesNo, "計算表印刷")
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
