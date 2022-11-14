VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCompany 
   Caption         =   "会社選択:年末調整のお知らせ作成"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "formCompany.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// フォーム起動時処理
Private Sub UserForm_Initialize()

    With Sheets("hidden")
        .Activate
        .Cells.ClearContents
    End With
    
    With Sheets("情報")
        .Range(.Cells(2, 2), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 2)).Copy
    End With
    
    With Sheets("hidden")
        .Cells(1, 1).PasteSpecial xlPasteValues
        .Columns(1).RemoveDuplicates (Array(1))
    End With
    
    cmbCompany.AddItem "全社"
    
    Dim i As Long
    
    For i = 1 To Sheets("hidden").Cells(Rows.Count, 1).End(xlUp).Row
        cmbCompany.AddItem Sheets("hidden").Cells(i, 1).value
    Next
        
    cmbCompany.Style = fmStyleDropDownList
    
    txtAddedCompany.MultiLine = True
    txtAddedCompany.Locked = True
    
End Sub

'// 実行するを押したときの処理(メインルーチン)
Private Sub cmdEnter_Click()

    Application.DisplayAlerts = False

    If txtAddedCompany.value = "" Then
        MsgBox "年調のお知らせを作成する会社を選択してください。", vbQuestion, "年調のお知らせ作成"
        Exit Sub
    End If
    
    '// リストに追加した会社を改行(vbLf)で分割したもの
    Dim splitedAddedCompany As Variant: splitedAddedCompany = Split(txtAddedCompany.value, vbCrLf)
    
    '// 改行(vbLf)で分割したものを配列に変換
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
    
    '// 年調資料作成
    If createYearEndAdjustmentNotice(selectedCompanies) Then
        MsgBox "処理が完了しました。" & vbLf & "デスクトップを確認してください。", Title:="年調のお知らせ作成"
    End If
    
    Unload formProgress
    Unload Me

End Sub

'// 「追加」を押したときの処理
Private Sub cmdAddCompany_Click()

    If cmbCompany.value = "" Then
        MsgBox "追加する会社名を選択してください。", vbQuestion, "年調お知らせ資料作成"
        Exit Sub
    End If
    
    txtAddedCompany.Locked = False
    
    
    If cmbCompany.value = "全社" Then
        Call addAllCompanies
    
    '// 年末のお知らせを作る会社リストが空白の場合
    ElseIf txtAddedCompany.value = "" Then
        txtAddedCompany.value = cmbCompany.value
    
    '// 年調のお知らせを作る会社リストに値があり、現在選択した会社がまだ追加されていない場合
    ElseIf InStr(1, txtAddedCompany.value, cmbCompany.value) = 0 Then
        txtAddedCompany.value = txtAddedCompany.value & vbLf & cmbCompany
    End If
    
    txtAddedCompany.Locked = True
    cmbCompany.value = ""
    
End Sub

'// 年調のお知らせを作る会社リストに全社追加
Private Sub addAllCompanies()

    Dim allCompanies As String
    Dim i As Long
        
    For i = 0 To cmbCompany.ListCount - 1
        If cmbCompany.List(i) = "全社" Then
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

'// リストをクリアを押したときの処理
Private Sub cmdClearCompany_Click()

    With txtAddedCompany
        .Locked = False
        .value = ""
        .Locked = True
    End With
    
End Sub

'// 閉じるを押したときの処理
Private Sub cmdClose_Click()

    Unload Me

End Sub
