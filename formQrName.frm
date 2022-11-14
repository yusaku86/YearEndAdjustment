VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formQrName 
   Caption         =   "QRコード名前変更"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895.001
   OleObjectBlob   =   "formQrName.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formQrName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// QRコード画像の名前を変更するフォーム
Option Explicit

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    Call addAllShapesName("btnChangeName")

    cmbSelectShape.Style = fmStyleDropDownList
    
    With cmbSelectCompany
        .AddItem "山岸運送"
        .AddItem "YCL"
        .AddItem "山岸ロジスターズ"
        .AddItem "東海YMトランス"
        
        .Style = fmStyleDropDownList
    End With
    
End Sub

'// シート「QRコード」にある全ての図形の名前をコンボボックスに追加
Private Sub addAllShapesName(ByVal exception As String)

    Dim tmpShape As Shape

    For Each tmpShape In Sheets("QRコード").Shapes
        If tmpShape.Name <> exception Then
            cmbSelectShape.AddItem tmpShape.Name
        End If
    Next

End Sub

'// 実行を押したときの処理:QRコード図形の名前を変更
Private Sub cmdEnter_Click()

    If cmbSelectShape.value = "" Then
        MsgBox "名前を変更するQRコードを選択してください。", vbQuestion, "年調お知らせ作成"
        Exit Sub
    ElseIf cmbSelectCompany.value = "" Then
        MsgBox "どの会社のQRコードかを選択してください。", vbQuestion, "年調お知らせ作成"
        Exit Sub
    End If

    Dim newName As String
    
    '// 選択した会社で名前を決定
    Select Case cmbSelectCompany.value
        Case "山岸運送":         newName = "YamagishiUnso"
        Case "YCL":              newName = "YCL"
        Case "山岸ロジスターズ": newName = "Logisters"
        Case "東海YMトランス":   newName = "Tokai"
    End Select

    Call changeNameOfShape(Sheets("QRコード").Shapes(cmbSelectShape.value), newName)
    
    cmbSelectShape.Clear
    cmbSelectCompany.value = ""
    
    Call addAllShapesName("btnChangeName")
    
    MsgBox "名前を変更しました。", Title:="年調お知らせ作成"

End Sub

'// 閉じるを押したときの処理
Private Sub cmdClose_Click()

    Unload Me

End Sub
