Attribute VB_Name = "M04_Check"
Option Explicit

Private cnA As New ADODB.Connection
Private rsA As New ADODB.Recordset
Private strKBN    As String  'x“X‹æ•ª
Private strDAT    As String  '¡‰ñx‹‹”NŒ
Private strDAZ    As String  '‘O‰ñx‹‹”NŒ
Private strDAL    As String  '‘O”Nx‹‹”NŒ

Const SQL1 = "SELECT * FROM Ü—^ WHERE (((x‹‹”NŒ) = '"
Const SQL2 = "') AND ((•”–å1)='"
Const SQL3 = "') And ((ĞˆõƒR[ƒh) = '"
Const SQL4 = "')) ORDER BY “™‹‰ DESC, ĞˆõƒR[ƒh"

Sub Proc_Check()

Dim lngR    As Long   '—ñ¶³İÀ

    '¼°Ä¸Ø±
    Range("A5:E153").ClearContents
    Range("H5:I153").ClearContents
    Range("L5:M153").ClearContents
    
    'x‹‹”NŒ¾¯Ä
    strDAT = Sheets("Main").Range("E2") & Format(Sheets("Main").Range("G2"), "00")
    If Right(strDAT, 2) = "12" Then
        '“~‹G ===============================================
        strDAZ = Sheets("Main").Range("E2") & "07"
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "12"
        Range("D3") = "¡‰ñ(“~‹Gj"
        Range("G3") = "‘O‰ñ(‰Ä‹Gj"
        Range("K3") = "‘O”N(“~‹Gj"
    Else
        '‰Ä‹G ===============================================
        strDAZ = CLng(Sheets("Main").Range("E2")) - 1 & "12"
        strDAL = CLng(Sheets("Main").Range("E2")) - 1 & "07"
        Range("D3") = "¡‰ñ(‰Ä‹Gj"
        Range("G3") = "‘O‰ñ(“~‹Gj"
        Range("K3") = "‘O”N(‰Ä‹Gj"
    End If
    '====================================================
    '–‹ÆŠ‹æ•ª‚²‚Æ“Ç‚İ
    strKBN = Range("S1")
    If strKBN = "" Then GoTo Exit_DB
    cnA.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbS
    cnA.Open
    strSQL = SQL1 & strDAT & SQL2 & strKBN & SQL4
    rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
    If rsA.EOF = False Then rsA.MoveFirst
    lngR = 5
    Do Until rsA.EOF
        Cells(lngR, 1) = strKBN
        Cells(lngR, 2) = rsA.Fields("ĞˆõƒR[ƒh")
        Cells(lngR, 3) = rsA.Fields("Ğˆõ–¼")
        Cells(lngR, 4) = rsA.Fields("’À‹à")
        Cells(lngR, 5) = rsA.Fields("Ü—^x‹‹Šz")
        rsA.MoveNext
        lngR = lngR + 1
    Loop
    
    rsA.Close
    lngR = 5
    Do
        If Cells(lngR, 2) = "" Then Exit Do
        strSQL = SQL1 & strDAZ & SQL2 & strKBN & SQL3 & Cells(lngR, 2) & SQL4
        rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsA.EOF = False Then
            rsA.MoveFirst
            Cells(lngR, 8) = rsA.Fields("’À‹à")
            Cells(lngR, 9) = rsA.Fields("Ü—^x‹‹Šz")
            rsA.MoveNext
        End If
        rsA.Close
        lngR = lngR + 1
    Loop
    
    lngR = 5
    Do
        If Cells(lngR, 2) = "" Then Exit Do
        strSQL = SQL1 & strDAL & SQL2 & strKBN & SQL3 & Cells(lngR, 2) & SQL4
        rsA.Open strSQL, cnA, adOpenStatic, adLockReadOnly
        If rsA.EOF = False Then
            rsA.MoveFirst
            Cells(lngR, 12) = rsA.Fields("’À‹à")
            Cells(lngR, 13) = rsA.Fields("Ü—^x‹‹Šz")
            rsA.MoveNext
        End If
        rsA.Close
        lngR = lngR + 1
    Loop
    
Exit_DB:
    cnA.Close

    Set rsA = Nothing
    Set cnA = Nothing

End Sub
