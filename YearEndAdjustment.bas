Attribute VB_Name = "YearEndAdjustment"
Option Explicit

'/**
 '* �N���̂��m�点�쐬(���C�����[�`��)
 '* @params targetCompanies() �N���̂��m�点���쐬������
'**/
Public Function createYearEndAdjustmentNotice(ByRef selectedCompanies() As Variant) As Boolean

    '// �����J�n�O�̊m�F
    Dim validationResult As String: validationResult = validate()
    
    If validationResult <> "passed" Then
        MsgBox validationResult, vbExclamation, "�N�����m�点�쐬"
        createYearEndAdjustmentNotice = False
        Exit Function
    End If

    '// �I�����ꂽ��Ђ̏]�ƈ��̍��v�l�����i���x���v�Z����ۂ̕���ƂȂ�
    Dim totalAmount As Long
    
    Dim i As Long
    For i = 0 To UBound(selectedCompanies)
        totalAmount = totalAmount + WorksheetFunction.CountIf(ThisWorkbook.Sheets("���").Columns(2), selectedCompanies(i))
    Next
    
    '// �I��������Ђ̐l�����V�[�g�ǉ� & �]�ƈ��ԍ��ύX
    Call addSheetForNumberOfEmployees(selectedCompanies, totalAmount)

    formPDF.Show vbModeless
    
    '// PDF�o��
    Dim companyName As String: companyName = createCompanyName(selectedCompanies)
    Call exportPDF(ActiveWorkbook, companyName)
    
    Unload formPDF
    
    ActiveWorkbook.Close False
    
    createYearEndAdjustmentNotice = True
    
End Function

'// �����J�n�O�̊m�F
Private Function validate() As String

    If existShape(Sheets("QR�R�[�h"), "YamagishiUnso") = False Or existShape(Sheets("QR�R�[�h"), "YCL") = False Or _
       existShape(Sheets("QR�R�[�h"), "Logisters") = False Or existShape(Sheets("QR�R�[�h"), "Tokai") = False Then
         
        validate = "�V�[�g�uQR�R�[�h�v�ɓK�؂Ȗ��O��QR�R�[�h�摜������܂���B"
        Exit Function
    ElseIf existShape(Sheets("���C�A�E�g"), "YamagishiUnso") = False And existShape(Sheets("���C�A�E�g"), "YCL") = False And _
           existShape(Sheets("���C�A�E�g"), "Logisters") = False And existShape(Sheets("���C�A�E�g"), "Tokai") = False Then
        
        validate = "�V�[�g�u���C�A�E�g�v�ɓK�؂Ȗ��O��QR�R�[�h�摜������܂���B"
        Exit Function
    End If
    
    validate = "passed"
    
End Function

'// �w��̃V�[�g�Ɏw��̐}�`�����݂��邩����
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
 '* �I��������Ђ̐l�����V�[�g��ǉ����A�Ј��ԍ��̕ύX��URL�EQR�R�[�h��K�؂Ȃ��̂ɕύX
 '*
 '* @params targetCompanies() �N���̂��m�点���쐬������
 '* @params totalAmount       �I��������Ђ̏]�ƈ��̐l��
 '**/
Private Sub addSheetForNumberOfEmployees(ByRef selectedCompanies() As Variant, ByVal totalAmount As Long)

    '// �V�[�g�u�ݒ�v�̎Ј��ԍ��EURL�EQR�R�[�h�Ɠ��͂���Ă���Z��
    Dim employeeIdCell As Range: Set employeeIdCell = Sheets("�ݒ�").Columns(1).Find(what:="�Ј��ԍ�", lookat:=xlWhole)
    Dim urlCell As Range:        Set urlCell = Sheets("�ݒ�").Columns(1).Find(what:="URL", lookat:=xlWhole)
    Dim qrCell As Range:         Set qrCell = Sheets("�ݒ�").Columns(1).Find(what:="QR�R�[�h", lookat:=xlWhole)

    '// ��L�̃Z������A�Ј��ԍ��EURL�EQR�R�[�h����͂���s�ԍ��E��ԍ����擾
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
    
    '// �i���x��\���t�H�[���N��
    formProgress.Show vbModeless

    Application.ScreenUpdating = False

    '// ����V�[�g
    Dim rosterSheet As Worksheet: Set rosterSheet = ThisWorkbook.Sheets("���")
    
    Dim lastRow As Long: lastRow = rosterSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long
        
    Dim processingCompany As String: processingCompany = rosterSheet.Cells(2, 2).value
    Dim formerProcessingCompany As String
    
    '// �V�[�g�ǉ��Ə]�ƈ��R�[�h����
    For i = 2 To lastRow

        '// ���݂̏]�ƈ��̏�����Ђ��Ώۂ̉�ЂɊ܂܂�Ă��Ȃ���Ύ��̃��[�v��
        If arrayIncludes(selectedCompanies, rosterSheet.Cells(i, 2).value) = False Then
            GoTo Continue
        End If
        
        '// �]�ƈ��R�[�h��������Ƃ��ĕۑ�����Ă���AVlookup�ŏE���Ȃ��Ȃ邽�ߐ��l�ɕϊ�
        rosterSheet.Cells(i, 1).value = Val(rosterSheet.Cells(i, 1).value)
        
        With newBook
        
            '// 2�l�ڈȍ~�͐V�����ǉ������u�b�N�̍Ō�̃V�[�g���R�s�[���ău�b�N�����ɒǉ�
            If .Sheets.Count >= 2 Then
                .Sheets(.Sheets.Count).Copy After:=.Sheets(.Sheets.Count)
            End If
        
            '// �ŏ��ɏ�������]�ƈ��͂��̃u�b�N�́u���C�A�E�g�v���R�s�[���ĐV�����ǉ������u�b�N�̖����ɒǉ�
            If .Sheets.Count = 1 Then
                '// QR�R�[�h���폜���āA�K�؂Ȃ��̂�\��t��
                On Error Resume Next
                
                With ThisWorkbook.Sheets("���C�A�E�g")
                    .Shapes("YamagishiUnso").Delete
                    .Shapes("YCL").Delete
                    .Shapes("Logisters").Delete
                    .Shapes("Tokai").Delete
                End With
                
                On Error GoTo 0
                
                Call pasteQR(processingCompany, ThisWorkbook.Sheets("���C�A�E�g"), qrRow, qrColumn)
                
                '// URL�̒l���������̉�Ђ̂��̂ɂ���
                Call setUrl(processingCompany, ThisWorkbook.Sheets("���C�A�E�g"), urlRow, urlColumn)
                
                '// �\��t����O�Ɉ���̐ݒ�
                Call setPrintMode(ThisWorkbook.Sheets("���C�A�E�g"))
                ThisWorkbook.Sheets("���C�A�E�g").Copy After:=.Sheets(1)
            End If
            
            '// �]�ƈ��ԍ������
            .Sheets(.Sheets.Count).Cells(employeeIdRow, employeeIdColumn).value = rosterSheet.Cells(i, 1).value
            
            '// �O�̏]�ƈ��̏�����Ђƌ��ݏ������Ă���]�ƈ��̏�����Ђ��قȂ�ꍇ
            If rosterSheet.Cells(i, 2).value <> processingCompany Then
                formerProcessingCompany = processingCompany
                processingCompany = rosterSheet.Cells(i, 2).value
                
                '// URL�̒l�ύX
                Call setUrl(processingCompany, .Sheets(.Sheets.Count), urlRow, urlColumn)
                '// �O�̏]�ƈ���QR�R�[�h�폜
                Call deleteFormerQr(formerProcessingCompany, .Sheets(.Sheets.Count))
                '// ���ݏ������Ă���Ј��̏�����Ђ�QR�R�[�h�摜��\��t��
                Call pasteQR(processingCompany, .Sheets(.Sheets.Count), qrRow, qrColumn)
            End If
            
            '// �V�[�g��10���ǉ����邲�Ƃɐi���x�\��
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
' * �z��ɒl���܂܂�Ă��邩����(������v��)
' * �������A�z��Ɋ܂܂��l��蕶�����������ꍇ��false�ɂȂ�
' * ���p�ƑS�p�͋�ʂ��Ȃ�
' *
' * ��)arrayIncludes(["�R�݉^����"],"�R�݉^��") �� true
' *    arrayIncludes(["�R�݉^��"],"�R�݉^����") �� false
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

'// URL�̒l��ύX
Private Sub setUrl(ByVal companyName As String, targetSheet As Worksheet, ByVal inputRow As Long, ByVal inputColumn As Long)
    
    Dim urlSheet As Worksheet: Set urlSheet = ThisWorkbook.Sheets("URL")
    Dim targetRow As Long
    
    If InStr(1, companyName, "�R�݉^��") > 0 Then
        targetRow = WorksheetFunction.Match("�R�݉^��", urlSheet.Columns(1), 0)
    ElseIf InStr(1, companyName, "YCL") > 0 Then
        targetRow = WorksheetFunction.Match("YCL", urlSheet.Columns(1), 0)
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "ۼ޽����") > 0 Then
        targetRow = WorksheetFunction.Match("�R�݃��W�X�^�[�Y", urlSheet.Columns(1), 0)
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "���CYM") > 0 Then
        targetRow = WorksheetFunction.Match("���CYM�g�����X", urlSheet.Columns(1), 0)
    End If

    targetSheet.Cells(inputRow, inputColumn).value = urlSheet.Cells(targetRow, 2).value
    
End Sub

'// ������Ђ�QR�R�[�h���R�s�[���ē\��t��
Private Sub pasteQR(ByVal companyName As String, ByVal destinationSheet As Worksheet, ByVal inputRow As Long, ByVal inputColumn As Long)
    
    Dim targetQr As Shape

    If InStr(1, companyName, "�R�݉^��") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QR�R�[�h").Shapes("YamagishiUnso")
    ElseIf InStr(1, companyName, "YCL") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QR�R�[�h").Shapes("YCL")
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "ۼ޽����") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QR�R�[�h").Shapes("Logisters")
    ElseIf InStr(1, StrConv(companyName, vbNarrow), "���CYM") > 0 Then
        Set targetQr = ThisWorkbook.Sheets("QR�R�[�h").Shapes("Tokai")
    End If
    
    targetQr.Copy
    
    '// QR�R�[�h��\��t���Ė��O��ύX(�\��t����Ɓu�}3�v�̂悤�Ȗ��O�ɂȂ邽��
    With destinationSheet
        .Cells(inputRow, inputColumn).PasteSpecial
        .Shapes(.Shapes.Count).Name = targetQr.Name
    End With
    
    Set targetQr = Nothing
End Sub

'// ������Ђ��O�̃��[�v�̏]�ƈ��ƈقȂ�ꍇ�A�O�̏]�ƈ��̏�����Ђ�QR�R�[�h�摜���폜
Private Sub deleteFormerQr(ByVal formerProcessingCompany As String, ByVal targetSheet As Worksheet)

    Dim targetQr As String
    
    If InStr(1, formerProcessingCompany, "�R�݉^��") > 0 Then
        targetQr = "YamagishiUnso"
    ElseIf InStr(1, formerProcessingCompany, "YCL") > 0 Then
        targetQr = "YCL"
    ElseIf InStr(1, StrConv(formerProcessingCompany, vbNarrow), "ۼ޽����") > 0 Then
        targetQr = "Logisters"
    ElseIf InStr(1, StrConv(formerProcessingCompany, vbNarrow), "���CYM") > 0 Then
        targetQr = "Tokai"
    End If
    
    On Error Resume Next
    
    targetSheet.Shapes(targetQr).Delete
    
    On Error GoTo 0
    
End Sub

'// �I�����ꂽ��Ж�����t�@�C�����Ŏg�p�����Ж����쐬
Private Function createCompanyName(ByRef selectedCompanies() As Variant) As String
    
    Dim companiesName As String
    Dim i As Long
    Dim selectedName As String
    
    For i = 0 To UBound(selectedCompanies)
        If InStr(1, selectedCompanies(i), "�R�݉^��") > 0 Then
            selectedName = "�R�݉^��"
        ElseIf InStr(1, selectedCompanies(i), "YCL") > 0 Then
            selectedName = "YCL"
        ElseIf InStr(1, StrConv(selectedCompanies(i), vbNarrow), "ۼ޽����") > 0 Then
            selectedName = "ۼ�"
        ElseIf InStr(1, StrConv(selectedCompanies(i), vbNarrow), "���CYM") > 0 Then
            selectedName = "���CYM"
        End If
    
        If companiesName = "" Then
            companiesName = selectedName
        Else
            companiesName = companiesName & "�E" & selectedName
        End If
    Next
    
    createCompanyName = companiesName

End Function

'// �S�V�[�g��1��PDF�t�@�C���Ƃ��ĕۑ�
Private Sub exportPDF(ByVal targetBook As Workbook, ByVal companiesName As String)
    
    Dim wsh As New WshShell
    
    targetBook.Sheets.Select
    targetBook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=wsh.SpecialFolders(4) & "\�N�����m�点(" & companiesName & ").pdf"
    
    Set wsh = Nothing
    
End Sub

'// �V�[�g�̈���ݒ�
'// �V�[�g��1���Ɏ��߂�A���������ɒ�����
Private Sub setPrintMode(ByVal targetSheet As Worksheet)

    With targetSheet.PageSetup
    
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .CenterHorizontally = True
    
    End With

End Sub

'// ��БI�����邽�߂̃t�H�[���N��
Public Sub openFormCompany()

    formCompany.Show

End Sub
