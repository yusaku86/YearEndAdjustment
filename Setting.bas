Attribute VB_Name = "Setting"
Option Explicit
'// 設定を行うモジュール

'// 図形の名前を変更
Public Sub changeNameOfShape(ByVal targetShape As Shape, ByVal newName As String)

    targetShape.Name = newName

End Sub

'// QRコード画像の名前を変更するためのフォーム起動
Public Sub openFormToChangeQrName()

    formQrName.Show

End Sub
