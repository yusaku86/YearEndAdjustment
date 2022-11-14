VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCompany 
   Caption         =   "��БI��:�N�������̂��m�点�쐬"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "formCompany.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// �t�H�[���N��������
Private Sub UserForm_Initialize()

    With Sheets("hidden")
        .Activate
        .Cells.ClearContents
    End With
    
    With Sheets("���")
        .Range(.Cells(2, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Copy
    End With
    
    With Sheets("hidden")
        .Cells(1, 1).PasteSpecial xlPasteValues
        .Columns(1).RemoveDuplicates (Array(1))
    End With
    
    cmbCompany.AddItem "�S��"
    
    Dim i As Long
    
    For i = 1 To Sheets("hidden").Cells(Rows.Count, 1).End(xlUp).Row
        cmbCompany.AddItem Sheets("hidden").Cells(i, 1).value
    Next
        
    cmbCompany.Style = fmStyleDropDownList
    
    txtAddedCompany.MultiLine = True
    txtAddedCompany.Locked = True
    
End Sub

'// ���s������������Ƃ��̏���(���C�����[�`��)
Private Sub cmdEnter_Click()

    Application.DisplayAlerts = False

    If txtAddedCompany.value = "" Then
        MsgBox "�N���̂��m�点���쐬�����Ђ�I�����Ă��������B", vbQuestion, "�N���̂��m�点�쐬"
        Exit Sub
    End If
    
    '// ���X�g�ɒǉ�������Ђ����s(vbLf)�ŕ�����������
    Dim splitedAddedCompany As Variant: splitedAddedCompany = Split(txtAddedCompany.value, vbCrLf)
    
    '// ���s(vbLf)�ŕ����������̂�z��ɕϊ�
    Dim selectedCompanies() As Variant
    ReDim selectedCompanies(0)
    
    Dim i As Long
    
    For i = 0 To UBound(splitedAddedCompany)
        If i <> 0 Then
            ReDim Preserve selectedCompanies(UBound(selectedCompanies) + 1)
        End If
        
        selectedCompanies(UBound(selectedCompanies)) = splitedAddedCompany(i)
    Next

    Me.Hide
    
    '// �N�������쐬
    If createYearEndAdjustmentNotice(selectedCompanies) Then
        MsgBox "�������������܂����B" & vbLf & "�f�X�N�g�b�v���m�F���Ă��������B", Title:="�N���̂��m�点�쐬"
    End If
    
    Unload formProgress
    Unload Me

End Sub

'// �u�ǉ��v���������Ƃ��̏���
Private Sub cmdAddCompany_Click()

    If cmbCompany.value = "" Then
        MsgBox "�ǉ������Ж���I�����Ă��������B", vbQuestion, "�N�����m�点�����쐬"
        Exit Sub
    End If
    
    txtAddedCompany.Locked = False
    
    
    If cmbCompany.value = "�S��" Then
        Call addAllCompanies
    
    '// �N���̂��m�点������Ѓ��X�g���󔒂̏ꍇ
    ElseIf txtAddedCompany.value = "" Then
        txtAddedCompany.value = cmbCompany.value
    
    '// �N���̂��m�点������Ѓ��X�g�ɒl������A���ݑI��������Ђ��܂��ǉ�����Ă��Ȃ��ꍇ
    ElseIf InStr(1, txtAddedCompany.value, cmbCompany.value) = 0 Then
        txtAddedCompany.value = txtAddedCompany.value & vbLf & cmbCompany
    End If
    
    txtAddedCompany.Locked = True
    cmbCompany.value = ""
    
End Sub

'// �N���̂��m�点������Ѓ��X�g�ɑS�Вǉ�
Private Sub addAllCompanies()

    Dim allCompanies As String
    Dim i As Long
        
    For i = 0 To cmbCompany.ListCount - 1
        If cmbCompany.List(i) = "�S��" Then
            GoTo Continue
        End If
        
        If allCompanies = "" Then
            allCompanies = cmbCompany.List(i)
        Else
            allCompanies = allCompanies & vbLf & cmbCompany.List(i)
        End If

Continue:
    Next
    
    txtAddedCompany.value = allCompanies

End Sub

'// ���X�g���N���A���������Ƃ��̏���
Private Sub cmdClearCompany_Click()

    With txtAddedCompany
        .Locked = False
        .value = ""
        .Locked = True
    End With
    
End Sub

'// ������������Ƃ��̏���
Private Sub cmdClose_Click()

    Unload Me

End Sub
