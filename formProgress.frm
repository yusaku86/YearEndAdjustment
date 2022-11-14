VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formProgress 
   Caption         =   "しばらくお待ちください。"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formProgress.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// 進捗度に合わせてプログレスバーの表示とラベルの表示変更
Public Sub indicateProgress(ByVal progress As Double)
    
    Application.ScreenUpdating = True
    
    Dim progressRate As Double: progressRate = progress * 100
    
    progressBar.value = Int(progressRate)
    
    If progressRate >= 100 Then
        progressRate = 100
    End If
    
    lblPercent.Caption = WorksheetFunction.Round(progressRate, 2) & "% 完了"

    Application.ScreenUpdating = False

End Sub

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    With progressBar
        .Min = 0
        .Max = 100
        .value = 0
    End With

End Sub
