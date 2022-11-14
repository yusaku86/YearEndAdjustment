VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formQrName 
   Caption         =   "QR�R�[�h���O�ύX"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895.001
   OleObjectBlob   =   "formQrName.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formQrName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// QR�R�[�h�摜�̖��O��ύX����t�H�[��
Option Explicit

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    Call addAllShapesName("btnChangeName")

    cmbSelectShape.Style = fmStyleDropDownList
    
    With cmbSelectCompany
        .AddItem "�R�݉^��"
        .AddItem "YCL"
        .AddItem "�R�݃��W�X�^�[�Y"
        .AddItem "���CYM�g�����X"
        
        .Style = fmStyleDropDownList
    End With
    
End Sub

'// �V�[�g�uQR�R�[�h�v�ɂ���S�Ă̐}�`�̖��O���R���{�{�b�N�X�ɒǉ�
Private Sub addAllShapesName(ByVal exception As String)

    Dim tmpShape As Shape

    For Each tmpShape In Sheets("QR�R�[�h").Shapes
        If tmpShape.Name <> exception Then
            cmbSelectShape.AddItem tmpShape.Name
        End If
    Next

End Sub

'// ���s���������Ƃ��̏���:QR�R�[�h�}�`�̖��O��ύX
Private Sub cmdEnter_Click()

    If cmbSelectShape.value = "" Then
        MsgBox "���O��ύX����QR�R�[�h��I�����Ă��������B", vbQuestion, "�N�����m�点�쐬"
        Exit Sub
    ElseIf cmbSelectCompany.value = "" Then
        MsgBox "�ǂ̉�Ђ�QR�R�[�h����I�����Ă��������B", vbQuestion, "�N�����m�点�쐬"
        Exit Sub
    End If

    Dim newName As String
    
    '// �I��������ЂŖ��O������
    Select Case cmbSelectCompany.value
        Case "�R�݉^��":         newName = "YamagishiUnso"
        Case "YCL":              newName = "YCL"
        Case "�R�݃��W�X�^�[�Y": newName = "Logisters"
        Case "���CYM�g�����X":   newName = "Tokai"
    End Select

    Call changeNameOfShape(Sheets("QR�R�[�h").Shapes(cmbSelectShape.value), newName)
    
    cmbSelectShape.Clear
    cmbSelectCompany.value = ""
    
    Call addAllShapesName("btnChangeName")
    
    MsgBox "���O��ύX���܂����B", Title:="�N�����m�点�쐬"

End Sub

'// ������������Ƃ��̏���
Private Sub cmdClose_Click()

    Unload Me

End Sub
