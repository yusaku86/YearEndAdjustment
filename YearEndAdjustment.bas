Attribute VB_Name = "YearEndAdjustment"
Option Explicit

'/**
 '* 年調のお知らせ作成(メインルーチン)
 '* @params targetCompanies() 年調のお知らせを作成する会社
'**/
Public Function createYearEndAdjustmentNotice(ByRef selectedCompanies() As Variant) As Boolean

    '// 処理開始前の確認
    Dim validationResult As String: validationResult = validate()
    
    If validationResult <> "passed" Then
        MsgBox validationResult, vbExclamation, "年調お知らせ作成"
        createYearEndAdjustmentNotice = False
        Exit Function
    End If

    '// 選択された会社の従業員の合計人数→進捗度を計算する際の分母となる
    Dim totalAmount As Long
    
    Dim i As Long
    For i = 0 To UBound(selectedCompanies)
        totalAmount = totalAmount + WorksheetFunction.CountIf(ThisWorkbook.Sheets("情報").Columns(2), selectedCompanies(i))
    Next
    
    '// 選択した会社の人数分シート追加 & 従業員番号変更
    Call addSheetForNumberOfEmployees(selectedCompanies, totalAmount)

    formPDF.Show vbModeless
    
    '// PDF出力
    Dim companyName As String: companyName = createCompanyName(selectedCompanies)
    Call exportPDF(ActiveWorkbook, companyName)
    
    Unload formPDF
    
    ActiveWorkbook.Close False
    
    createYearEndAdjustmentNotice = True
    
End Function

'// 処理開始前の確認
Private Function validate() As String

    If existShape(Sheets("QRコード"), "YamagishiUnso") = False Or existShape(Sheets("QRコード"), "YCL") = False Or _
       existShape(Sheets("QRコード"), "Logisters") = False Or existShape(Sheets("QRコード"), "Tokai") = False Then
         
        validate = "シート「QRコード」に適切な名前のQRコード画像がありません。"
        Exit Function
    ElseIf existShape(Sheets("レイアウト"), "YamagishiUnso") = False And existShape(Sheets("レイアウト"), "YCL") = False And _
           existShape(Sheets("レイアウト"), "Logisters") = False And existShape(Sheets("レイアウト"), "Tokai") = False Then
        
        validate = "シート「レイアウト」に適切な名前のQRコード画像がありません。"
        Exit Function
    End If
    
    validate = "passed"
    
End Function

'// 指定のシートに指定の図形が存在するか判定
Private Function existShape(ByVal targetSheet As Worksheet, ByVal shapeName As String) As Boolean

    Dim tmpShape As Shape
    
    For Each tmpShape In targetSheet.Shapes
        If tmpShape.Name = shapeName Then
            existShape = True
            Exit Function
        End If
    Next
    
    existShape = False

End Function

'/**
 '* 選択した会社の人数分シートを追加し、社員番号の変更とURL・QRコードを適切なものに変更
 '*
 '* @params targetCompanies() 年調のお知らせを作成する会社
 '* @params totalAmount       選択した会社の従業員の人数
 '**/
Private Sub addSheetForNumberOfEmployees(ByRef selectedCompanies() As Variant, ByVal totalAmount As Long)

    '// シート「設定」の社員番号・URL・QRコードと入力されているセル
    Dim employeeIdCell As Range: Set employeeIdCell = Sheets("設定").Columns(1).Find(what:="社員番号", lookat:=xlWhole)
    Dim urlCell As Range:        Set urlCell = Sheets("設定").Columns(1).Find(what:="URL", lookat:=xlWhole)
    Dim qrCell As Range:         Set qrCell = Sheets("設定").Columns(1).Find(what:="QRコード", lookat:=xlWhole)

    '// 上記のセルから、社員番号・URL・QRコードを入力する行番号・列番号を取得
    Dim employeeIdRow As Long:    employeeIdRow = employeeIdCell.Offset(, 1).value
    Dim employeeIdColumn As Long: employeeIdColumn = employeeIdCell.Offset(, 2).value
    
    Dim urlRow As Long:    urlRow = urlCell.Offset(, 1).value
    Dim urlColumn As Long: urlColumn = urlCell.Offset(, 2).value
    
    Dim qrRow As Long:    qrRow = qrCell.Offset(, 1).value
    Dim qrColumn As Long: qrColumn = qrCell.Offset(, 2).value
    
    Set employeeIdCell = Nothing
    Set urlCell = Nothing
    Set qrCell = Nothing
    
    Dim newBook As Workbook: Set newBook = Workbooks.Add
    
    '// 進捗度を表すフォーム起動
    formProgress.Show vbModeless

    Application.ScreenUpdating = False

    '// 名簿シート
    Dim rosterSheet As Worksheet: Set rosterSheet = ThisWorkbook.Sheets("情報")
    
    Dim lastRow As Long: lastRow = rosterSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long
        
    Dim processingCompany As String: processingCompany = rosterSheet.Cells(2, 2).value
    Dim formerProcessingCompany As String
    
    '// シート追加と従業員コード入力
    For i = 2 To lastRow

        '// 現在の従業員の所属会社が対象の会社に含まれていなければ次のループへ
        If arrayIncludes(selectedCompanies, rosterSheet.Cells(i, 2).value) = False Then
            GoTo Continue
        End If
        
        '// 従業員コードが文字列として保存されており、Vlookupで拾えなくなるため数値に変換
        rosterSheet.Cells(i, 1).value = Val(rosterSheet.Cells(i, 1).value)
        
        With newBook
        
            '// 2人目以降は新しく追加したブックの最後のシートをコピーしてブック末尾に追加
            If .Sheets.Count >= 2 Then
                .Sheets(.Sheets.Count).Copy After:=.Sheets(.Sheets.Count)
            End If
        
            '// 最初に処理する従業員はこのブックの「レイアウト」をコピーして新しく追加したブックの末尾に追加
            If .Sheets.Count = 1 Then
                '// QRコードを削除して、適切なものを貼り付け
                On Error Resume Next
                
                With ThisWorkbook.Sheets("レイアウト")
                    .Shapes("YamagishiUnso").Delete
                    .Shapes("YCL").Delete
                    .Shapes("Logisters").Delete
                    .Shapes("Tokai").Delete
                End With
                
                On Error GoTo 0
                
                Call pasteQR(processingCompany, ThisWorkbook.Sheets("レイアウト"), qrRow, qrColumn)
                
                '// URLの値を処理中の会社のものにする
                Call setUrl(processingCompany, ThisWorkbook.Sheets("レイアウト"), urlRow, urlColumn)
                
                '// 貼り付ける前に印刷の設定
                Call setPrintMode(ThisWorkbook.Sheets("レイアウト"))
                ThisWorkbook.Sheets("レイアウト").Copy After:=.Sheets(1)
            End If
            
            '// 従業員番号を入力
            .Sheets(.Sheets.Count).Cells(employeeIdRow, employeeIdColumn).value = rosterSheet.Cells(i, 1).value
            
            '// 前の従業員の所属会社と現在処理している従業員の所属会社が異なる場合
            If rosterSheet.Cells(i, 2).value <> processingCompany Then
                formerProcessingCompany = processingCompany
                processingCompany = rosterSheet.Cells(i, 2).value
                
                '// URLの値変更
                Call setUrl(processingCompany, .Sheets(.Sheets.Count), urlRow, urlColumn)
                '// 前の従業員のQRコード削除
                Call deleteFormerQr(formerProcessingCompany, .Sheets(.Sheets.Count))
                '// 現在処理している社員の所属会社のQRコード画像を貼り付け
                Call pasteQR(processingCompany, .Sheets(.Sheets.Count), qrRow, qrColumn)
            End If
            
            '// シートを10枚追加するごとに進捗度表示
            If .Sheets.Count Mod 10 = 0 Then
                formProgress.indicateProgress .Sheets.Count / totalAmount
            End If
        End With
Continue:
    Next
    
    newBook.Sheets(1).Delete
        
    Set newBook = Nothing
    
    Unload formProgress
        
End Sub

'/**
' * 配列に値が含まれているか判定(部分一致可)
' * ただし、配列に含まれる値より文字数が長い場合はfalseになる
' * 半角と全角は区別しない
' *
' * 例)arrayIncludes(["山岸運送㈱"],"山岸運送") は true
' *    arrayIncludes(["山岸運送"],"山岸運送㈱") は false
' **/
Private Function arrayIncludes(ByRef targetArray() As Variant, ByVal value As Variant) As Boolean

    Dim i As Long
    
    For i = 0 To UBound(targetArray)
        If InStr(1, StrConv(targetArray(i), vbNarrow), StrConv(value, vbNarrow)) > 0 Then
            arrayIncludes = True
            Exit Function
        End If
    Next
    
    arrayIncludes = False
    
End Function

'// URLの値を変更
Private Sub setUrl(ByVal companyName As String, targetSheet As Worksheet, ByVal inputRow As Long, ByVal inputColumn As Long)
    
    Dim urlSheet As Worksheet: Set urlSheet = ThisWorkbook.Sheets("URL")
    Dim targetRow As Long
    
    If InStr(1, companyName, "山岸運送") > 0 Then
        targetRow = WorksheetFunction.Match("山岸運送", urlSheet.Columns(1), 0)
    ElseIf InStr(1, companyName, "YCL") > 0 Then
        targetRow = WorksheetFunction.Match("YCL", urlSheet.Columns(1), 0)
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "ﾛｼﾞｽﾀｰｽﾞ") > 0 Then
        targetRow = WorksheetFunction.Match("山岸ロジスターズ", urlSheet.Columns(1), 0)
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "東海YM") > 0 Then
        targetRow = WorksheetFunction.Match("東海YMトランス", urlSheet.Columns(1), 0)
    End If

    targetSheet.Cells(inputRow, inputColumn).value = urlSheet.Cells(targetRow, 2).value
    
End Sub

'// 所属会社のQRコードをコピーして貼り付け
Private Sub pasteQR(ByVal companyName As String, ByVal destinationSheet As Worksheet, ByVal inputRow As Long, ByVal inputColumn As Long)
    
    Dim targetQr As Shape

    If InStr(1, companyName, "山岸運送") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QRコード").Shapes("YamagishiUnso")
    ElseIf InStr(1, companyName, "YCL") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QRコード").Shapes("YCL")
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "ﾛｼﾞｽﾀｰｽﾞ") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QRコード").Shapes("Logisters")
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "東海YM") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QRコード").Shapes("Tokai")
    End If
    
    targetQr.Copy
    
    '// QRコードを貼り付けて名前を変更(貼り付けると「図3」のような名前になるため
    With destinationSheet
        .Cells(inputRow, inputColumn).PasteSpecial
        .Shapes(.Shapes.Count).Name = targetQr.Name
    End With
    
    Set targetQr = Nothing
End Sub

'// 所属会社が前のループの従業員と異なる場合、前の従業員の所属会社のQRコード画像を削除
Private Sub deleteFormerQr(ByVal formerProcessingCompany As String, ByVal targetSheet As Worksheet)

    Dim targetQr As String
    
    If InStr(1, formerProcessingCompany, "山岸運送") > 0 Then
        targetQr = "YamagishiUnso"
    ElseIf InStr(1, formerProcessingCompany, "YCL") > 0 Then
        targetQr = "YCL"
    ElseIf InStr(1, StrConv(formerProcessingCompany, vbNarrow), "ﾛｼﾞｽﾀｰｽﾞ") > 0 Then
        targetQr = "Logisters"
    ElseIf InStr(1, StrConv(formerProcessingCompany, vbNarrow), "東海YM") > 0 Then
        targetQr = "Tokai"
    End If
    
    On Error Resume Next
    
    targetSheet.Shapes(targetQr).Delete
    
    On Error GoTo 0
    
End Sub

'// 選択された会社名からファイル名で使用する会社名を作成
Private Function createCompanyName(ByRef selectedCompanies() As Variant) As String
    
    Dim companiesName As String
    Dim i As Long
    Dim selectedName As String
    
    For i = 0 To UBound(selectedCompanies)
        If InStr(1, selectedCompanies(i), "山岸運送") > 0 Then
            selectedName = "山岸運送"
        ElseIf InStr(1, selectedCompanies(i), "YCL") > 0 Then
            selectedName = "YCL"
        ElseIf InStr(1, StrConv(selectedCompanies(i), vbNarrow), "ﾛｼﾞｽﾀｰｽﾞ") > 0 Then
            selectedName = "ﾛｼﾞ"
        ElseIf InStr(1, StrConv(selectedCompanies(i), vbNarrow), "東海YM") > 0 Then
            selectedName = "東海YM"
        End If
    
        If companiesName = "" Then
            companiesName = selectedName
        Else
            companiesName = companiesName & "・" & selectedName
        End If
    Next
    
    createCompanyName = companiesName

End Function

'// 全シートを1つのPDFファイルとして保存
Private Sub exportPDF(ByVal targetBook As Workbook, ByVal companiesName As String)
    
    Dim wsh As New WshShell
    
    targetBook.Sheets.Select
    targetBook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=wsh.SpecialFolders(4) & "\年調お知らせ(" & companiesName & ").pdf"
    
    Set wsh = Nothing
    
End Sub

'// シートの印刷設定
'// シートを1枚に収める、水平方向に中央寄せ
Private Sub setPrintMode(ByVal targetSheet As Worksheet)

    With targetSheet.PageSetup
    
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .CenterHorizontally = True
    
    End With

End Sub

'// 会社選択するためのフォーム起動
Public Sub openFormCompany()

    formCompany.Show

End Sub
