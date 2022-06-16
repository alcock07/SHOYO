Attribute VB_Name = "M01_Masta"
Option Explicit

Public Const dbM As String = "\\192.168.128.4\hb\kyuyo\グループ賃金.accdb"
Public Const dbT As String = "\\192.168.128.4\hb\ta\給与システム\グループ賃金.accdb"
Public Const dbS As String = "\\192.168.128.4\hb\kyuyo\賞与\賞与データ.accdb"
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

Dim cnA As New ADODB.Connection
Dim rsA As New ADODB.Recordset
Dim strKBN As String
Dim lngR   As Long
Dim lngC   As Long
Dim DateA  As Date
Dim DateB  As Date
Dim strYY  As String
Dim lngMM  As Long

    'ｼｰﾄ初期化
    Range("A4:J152").ClearContents
    Range("L4:L52").ClearContents
    Range("N4:N52").ClearContents
    
    '拠点区分判定して接続DB切替え
    strKBN = Range("Q2")
    If strKBN = "" Then GoTo Exit_DB
    If strKBN = "TA" Or strKBN = "KA" Then
        strDB = dbT  '\\192.168.128.4\hb\ta\給与システム\グループ賃金.accdb
    Else
        strDB = dbM  '\\192.168.128.4\hb\kyuyo\グループ賃金.accdb
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    
    '事業所区分ごと読込み
    strSQL = ""
    strSQL = strSQL & "SELECT 事業所区分,"
    strSQL = strSQL & "       社員コード,"
    strSQL = strSQL & "       社員名,"
    strSQL = strSQL & "       社員種類,"
    strSQL = strSQL & "       等級,"
    strSQL = strSQL & "       基本給１,"
    strSQL = strSQL & "       基本給２,"
    strSQL = strSQL & "       管理職手当,"
    strSQL = strSQL & "       家族手当,"
    strSQL = strSQL & "       部門1,"
    strSQL = strSQL & "       部門2,"
    strSQL = strSQL & "       部門3,"
    strSQL = strSQL & "       部門名,"
    strSQL = strSQL & "       入社年月日"
    strSQL = strSQL & "     FROM グループ社員マスター"
    strSQL = strSQL & "          WHERE 事業所区分 = '" & strKBN & "'"
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 4
    Do Until rsA.EOF
        '各項目ｾｯﾄ
        For lngC = 0 To 8
            Cells(lngR, lngC + 1) = rsA.Fields(lngC)
        Next lngC
        '部門区分ｾｯﾄ
        If IsNull(rsA.Fields("部門2")) = False Then Cells(lngR, 10) = rsA.Fields("部門2")
        If IsNull(rsA.Fields("部門3")) = False Then Cells(lngR, 12) = rsA.Fields("部門3")
        '生年月日
        If rsA.Fields("入社年月日") <> "" Then
            DateA = rsA.Fields("入社年月日")
        End If
        '新入社員判定処理
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
                Cells(lngR, 14) = "○"
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
Dim strKBN As String
Dim strCD  As String
Dim strKB1 As String
Dim strKB2 As String
Dim strKB3 As String
Dim lngR   As Long
Dim lngC   As Long
    
    '拠点区分判定して接続DB切替え
    strKBN = Range("Q2")
    If strKBN = "" Then GoTo Exit_DB
    If strKBN = "TA" Or strKBN = "KA" Then
        strDB = dbT  '\\192.168.128.4\hb\ta\給与システム\グループ賃金.accdb
    Else
        strDB = dbM  '\\192.168.128.4\hb\kyuyo\グループ賃金.accdb
    End If
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnA.Open
    
    lngR = 4
    Do
        strCD = Cells(lngR, 2) '社員ｺｰﾄﾞ
        If strCD = "" Then Exit Do
        strKB1 = Range("P2")
        strKB2 = Cells(lngR, 10)
        strKB3 = Cells(lngR, 12)
        If strCD <> "" Then
            'ﾏｽﾀ呼出
            strSQL = ""
            strSQL = strSQL & "SELECT 部門1,"
            strSQL = strSQL & "       部門2,"
            strSQL = strSQL & "       部門3,"
            strSQL = strSQL & "       部門名,"
            strSQL = strSQL & "       新入社員"
            strSQL = strSQL & "     FROM グループ社員マスター"
            strSQL = strSQL & "          WHERE 社員コード = '" & strCD & "'"
            rsA.Open strSQL, cnA, adOpenStatic, adLockPessimistic
            If rsA.EOF = False Then
                rsA.MoveFirst
                rsA.Fields(0) = strKB1
                rsA.Fields(1) = strKB2
                rsA.Fields(2) = strKB3
                rsA.Fields(3) = Cells(lngR, 13)
                If Cells(lngR, 14) = "○" Then
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
    
    MsgBox "登録しました(^^♪", vbInformation, "マスタ登録"
    
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
