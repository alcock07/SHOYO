Attribute VB_Name = "M00_Main"
Option Explicit

'コンピューター名を取得する関数の宣言
#If VBA7 Then
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public Const MAX_COMPUTERNAME_LENGTH = 15

Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' コンピューター名の長さを設定
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' コンピューター名を取得
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
                                            lngComputerNameLength)
    ' コンピューター名を表示
    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)
    
End Function

Sub Back_Main()
    Sheets("Main").Select
End Sub

Sub Back_Masta()
    Sheets("Masta").Select
End Sub

Sub Back_Data()
    Sheets("Data").Select
End Sub

Sub AP_END()
'==================
' 終了処理　Ver2.0
'==================

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
    Application.ReferenceStyle = xlA1
    Application.MoveAfterReturnDirection = xlDown
    Application.DisplayAlerts = False
    
    ThisWorkbook.Save
    strFN = ThisWorkbook.Name 'このブックの名前
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  'ファイルを閉じる
    Else
        Application.Quit  'Excellを終了
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub
