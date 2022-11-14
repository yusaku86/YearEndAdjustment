VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formProgress 
   Caption         =   "���΂炭���҂����������B"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formProgress.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// �i���x�ɍ��킹�ăv���O���X�o�[�̕\���ƃ��x���̕\���ύX
Public Sub indicateProgress(ByVal progress As Double)
    
    Application.ScreenUpdating = True
    
    Dim progressRate As Double: progressRate = progress * 100
    
    progressBar.value = Int(progressRate)
    
    If progressRate >= 100 Then
        progressRate = 100
    End If
    
    lblPercent.Caption = WorksheetFunction.Round(progressRate, 2) & "% ����"

    Application.ScreenUpdating = False

End Sub

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    With progressBar
        .Min = 0
        .Max = 100
        .value = 0
    End With

End Sub
