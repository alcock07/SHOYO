Attribute VB_Name = "M05_Personal"
Option Explicit

Sub 個別印刷()
Dim strName   As String
Dim lngR      As Long
Dim strKYU    As String
Dim lngKIN(6) As Long
Dim dblRT(3)  As Double
Dim strKHN    As String
Dim dblKHN    As Double
Dim lngKHN    As Long
Dim strNEW    As String
Dim vMSG      As Variant
    lngR = 7
    Do
        Sheets("Data").Select
        strName = Cells(lngR, 3)
        strKYU = Cells(lngR, 4)
        strKHN = Cells(lngR, 6)
        If InStr(1, strName, "総合計") <> 0 Then Exit Do
        If InStr(1, strKHN, "基本") <> 0 Then
            lngKHN = (InStr(1, strKHN, "(") + 1)
            If Len(strKHN) = 8 Then
                dblKHN = CDbl(Mid(strKHN, lngKHN, 4))
            ElseIf Len(strKHN) = 7 Then
                dblKHN = CDbl(Mid(strKHN, lngKHN, 3))
            ElseIf Len(strKHN) = 5 Then
                dblKHN = CDbl(Mid(strKHN, lngKHN, 1))
            Else
                MsgBox "？？？"
            End If
        End If
        strNEW = Cells(lngR, 26)
        If strKYU <> "" Then
            
            lngKIN(0) = Cells(lngR, 5)
            lngKIN(1) = Cells(lngR, 6)
            lngKIN(2) = Cells(lngR, 8)
            lngKIN(3) = Cells(lngR, 12)
            lngKIN(4) = Cells(lngR, 13)
            lngKIN(5) = Cells(lngR, 16)
            lngKIN(6) = Cells(lngR, 17)
            dblRT(0) = Cells(lngR, 7)
            dblRT(1) = Cells(lngR, 14)
            dblRT(2) = Cells(lngR, 15)
            dblRT(3) = Cells(lngR, 18)
            
            Sheets("PD").Select
            Cells(5, 3) = strName
            Cells(7, 3) = strKYU
            Cells(8, 3) = lngKIN(0)
            If strNEW = "Y" Then
                Cells(9, 3) = 0
                Cells(10, 3) = 0
                Cells(11, 3) = 0
                Cells(12, 3) = 0
                Cells(13, 3) = 0
            Else
                Cells(9, 3) = dblKHN
                Cells(10, 3) = lngKIN(1)
                Cells(11, 3) = dblRT(0)
                Cells(12, 3) = lngKIN(2)
                Cells(13, 3) = lngKIN(3)
            End If
            Cells(14, 3) = lngKIN(4)
            Cells(15, 3) = dblRT(1)
            Cells(16, 3) = dblRT(2)
            Cells(18, 3) = lngKIN(5)
            Cells(19, 3) = lngKIN(6)
            Cells(20, 3) = dblRT(3)
            
'            vMSG = MsgBox("印刷しますか？ " & strName, vbYesNoCancel, "印刷選択")
'            If vMSG = vbCancel Then Exit Do
'            If vMSG = vbYes Then
            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
            ActiveSheet.DisplayPageBreaks = False
'            End If
        End If
        
        lngR = lngR + 1
    Loop
    
    Sheets("Data").Select
    
End Sub

Sub 個別印刷R()

Dim strName   As String
Dim lngR      As Long
Dim strKYU    As String
Dim lngKIN(6) As Long
Dim dblRT(3)  As Double
Dim strKHN    As String
Dim dblKHN    As Double
Dim lngKHN    As Long
Dim strPP     As String

    lngR = 7
    Do
        Sheets("Data").Select
        strName = Cells(lngR, 3)
        strKYU = Cells(lngR, 4)
        strKHN = Cells(lngR, 6)
        strPP = Cells(lngR, 25)
        If InStr(1, strName, "総合計") <> 0 Then Exit Do
        If InStr(1, strKHN, "基本") <> 0 Then
            lngKHN = (InStr(1, strKHN, "(") + 1)
            If Len(strKHN) = 8 Then
                dblKHN = CDbl(Mid(strKHN, lngKHN, 4))
            ElseIf Len(strKHN) = 7 Then
                dblKHN = CDbl(Mid(strKHN, lngKHN, 3))
            ElseIf Len(strKHN) = 5 Then
                dblKHN = CDbl(Mid(strKHN, lngKHN, 1))
            Else
                MsgBox "？？？"
            End If
        End If
        
        If strKYU <> "" Then
        lngKIN(0) = Cells(lngR, 5)
        lngKIN(1) = Cells(lngR, 6)
        lngKIN(2) = Cells(lngR, 8)
        lngKIN(3) = Cells(lngR, 12)
        lngKIN(4) = Cells(lngR, 13)
        lngKIN(5) = Cells(lngR, 16)
        lngKIN(6) = Cells(lngR, 17)
        dblRT(0) = Cells(lngR, 7)
        dblRT(1) = Cells(lngR, 14)
        dblRT(2) = Cells(lngR, 15)
        dblRT(3) = Cells(lngR, 18)
        
        Sheets("PDR").Select
        Cells(2, 1) = strName & "  殿"
        Cells(16, 3) = lngKIN(4)
'        If strPP = "P" Then
'            Cells(15, 2) = ""
'        Else
'            Cells(15, 2) = "*給与一ヵ月相当"
'        End If
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        ActiveSheet.DisplayPageBreaks = False
        
        End If
        
        lngR = lngR + 1
    Loop
    
End Sub
