Attribute VB_Name = "Setting"
Option Explicit
'// �ݒ���s�����W���[��

'// �}�`�̖��O��ύX
Public Sub changeNameOfShape(ByVal targetShape As Shape, ByVal newName As String)

    targetShape.Name = newName

End Sub

'// QR�R�[�h�摜�̖��O��ύX���邽�߂̃t�H�[���N��
Public Sub openFormToChangeQrName()

    formQrName.Show

End Sub
