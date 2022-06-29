Attribute VB_Name = "M04_Check"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private strKBN    As String  '支店区分
Private strDAT    As String  '今回支給年月
Private strDAZ    As String  '前回支給年月
Private strDAL    As String  '前年支給年月

'Const SQL1 = "SELECT * FROM 賞与 WHERE (((支給年月) = '"
'Const SQL2 = "') AND ((部門1)='"
'Const SQL3 = "') And ((社員コード) = '"
'Const SQL4 = "')) ORDER BY 等級 DESC, 社員コード"

Sub Proc_Check()

Dim lngR    As Long   '列ｶｳﾝﾀ

    'ｼｰﾄｸﾘｱ
    Range("A5:E153").ClearContents
    Range("H5:I153").ClearContents
    Range("L5:M153").ClearContents
    
    '支給年月ｾｯﾄ
    strDAT = Sheets("Main").Range("E2") & Format(Sheets("Main").Range("G2"), "00")
    If Right(strDAT, 2) = "12" Then
        '冬季 ===============================================
        strDAZ = Sheets("Main").Range("E2") & "07"
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "12"
        Range("D3") = "今回(冬季）"
        Range("G3") = "前回(夏季）"
        Range("K3") = "前年(冬季）"
    Else
        '夏季 ===============================================
        strDAZ = CLng(Sheets("Main").Range("E2")) - 1 & "12"
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "07"
        Range("D3") = "今回(夏季）"
        Range("G3") = "前回(冬季）"
        Range("K3") = "前年(夏季）"
    End If
    '====================================================
    '事業所区分ごと読込み
    strKBN = Range("S1")
    If strKBN = "" Then GoTo Exit_DB
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open
'    strSQL = SQL1 & strDAT & SQL2 & strKBN & SQL4
    strSQL = ""
    strSQL = strSQL & "SELECT *"
    strSQL = strSQL & "       FROM 賞与"
    strSQL = strSQL & "            WHERE 支給年月 = '" & strDAT & "'"
    strSQL = strSQL & "            AND 部門1 = '" & strKBN & "'"
    strSQL = strSQL & "       ORDER BY 等級 DESC"
    strSQL = strSQL & "                ,社員コード"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 5
    Do Until rsA.EOF
        Cells(lngR, 1) = strKBN
        Cells(lngR, 2) = rsA.Fields("社員コード")
        Cells(lngR, 3) = rsA.Fields("社員名")
        Cells(lngR, 4) = rsA.Fields("賃金")
        Cells(lngR, 5) = rsA.Fields("賞与支給額")
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
        strSQL = strSQL & "       FROM 賞与"
        strSQL = strSQL & "            WHERE 支給年月 = '" & strDAZ & "'"
        strSQL = strSQL & "            AND 部門1 = '" & strKBN & "'"
        strSQL = strSQL & "            AND 社員コード = '" & Cells(lngR, 2) & "'"
        strSQL = strSQL & "       ORDER BY 等級 DESC"
        strSQL = strSQL & "              , 社員コード"
        rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsA.EOF = False Then
            rsA.MoveFirst
            Cells(lngR, 8) = rsA.Fields("賃金")
            Cells(lngR, 9) = rsA.Fields("賞与支給額")
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
        strSQL = strSQL & "       FROM 賞与"
        strSQL = strSQL & "            WHERE 支給年月 = '" & strDAL & "'"
        strSQL = strSQL & "            AND 部門1 = '" & strKBN & "'"
        strSQL = strSQL & "            AND 社員コード = '" & Cells(lngR, 2) & "'"
        strSQL = strSQL & "       ORDER BY 等級 DESC"
        strSQL = strSQL & "              , 社員コード"
        rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsA.EOF = False Then
            rsA.MoveFirst
            Cells(lngR, 12) = rsA.Fields("賃金")
            Cells(lngR, 13) = rsA.Fields("賞与支給額")
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
