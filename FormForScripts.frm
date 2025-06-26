VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormForScripts 
   Caption         =   "Script Form"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   OleObjectBlob   =   "FormForScripts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormForScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
FormForScripts.hide
End Sub
Private Sub StartButton_Click()
If FormForScripts.PathTB.Text = "" Then
    MsgBox "введите путь и попробуйте еще раз"
    Exit Sub
Else
If dxfcheck = True And pdfcheck = True And ReSaveCheck = True And updatecheck = True And excelcheck = True And ExcelCreateCheck = True Then
    MsgBox "ТЯЖЕЛО"
End If
If dxfcheck = False And pdfcheck = False And ReSaveCheck = False And updatecheck = False And excelcheck = False And ExcelCreateCheck = False Then
    MsgBox "А что делать то?"
End If
    If dxfcheck = True Then
        DXFForAllDocuments
    End If
    If pdfcheck = True Then
        ConvertDrawingsToPDF
    End If
    If updatecheck = True Then
        UpdateFileProperties
    End If
    If ReSaveCheck = True Then
        ReSaveFile
    End If
    If excelcheck = True Then
         ConvertAllExcelsToPDF
    End If
    If ExcelCreateCheck = True Then
        CreateExcelSpecification
    End If
End If
End Sub
Private Sub ConvertDrawingsToPDF()
ThisApplication.SilentOperation = True
    Dim folderPath As String
    folderPath = FormForScripts.PathTB.Text
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim rulePath As String
    rulePath = "C:\\Users\\project15\\Documents\\ilogicRules\\ConvertToPdf.iLogicVb"

    Dim recursive As Boolean
    recursive = FormForScripts.OBallfolders.Value  ' True если выбрана опция "все папки"

    Dim invApp As Application
    Set invApp = GetObject(, "Inventor.Application")

    Dim iLogicAuto As Object
    Set iLogicAuto = GetiLogicAddin(invApp)

    Dim exts As Variant
    exts = Array("idw", "dwg")

    ProcessFiles folderPath, recursive, exts, rulePath, invApp, iLogicAuto
ThisApplication.SilentOperation = False
    MsgBox "Преобразование в PDF завершено!"
End Sub


Private Sub DXFForAllDocuments()
ThisApplication.SilentOperation = True
    Dim folderPath As String
    folderPath = FormForScripts.PathTB.Text
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim rulePath As String
    rulePath = "C:\\Users\\project15\\Documents\\ilogicRules\\dxf.iLogicVb"

    Dim recursive As Boolean
    recursive = FormForScripts.OBallfolders.Value

    Dim invApp As Application
    Set invApp = GetObject(, "Inventor.Application")

    Dim iLogicAuto As Object
    Set iLogicAuto = GetiLogicAddin(invApp)

    Dim exts As Variant
    exts = Array("ipt", "iam")

    ProcessFiles folderPath, recursive, exts, rulePath, invApp, iLogicAuto
ThisApplication.SilentOperation = False
    MsgBox "Преобразование в DXF завершено!"
End Sub
Private Sub UpdateFileProperties()
ThisApplication.SilentOperation = True
    Dim folderPath As String
    folderPath = FormForScripts.PathTB.Text
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim rulePath As String
    rulePath = "C:\\Users\\project15\\Documents\\ilogicRules\\ChangePropertiesByFileName.iLogicVb"

    Dim recursive As Boolean
    recursive = FormForScripts.OBallfolders.Value  ' True если выбрана опция "все папки"

    Dim invApp As Application
    Set invApp = GetObject(, "Inventor.Application")

    Dim iLogicAuto As Object
    Set iLogicAuto = GetiLogicAddin(invApp)

    Dim exts As Variant
    If FormForScripts.UpdatePart.Value = True And FormForScripts.UpdateAssembly.Value = True And FormForScripts.UpdateDraft.Value = False Then
     ' Только детали и сборки
        exts = Array("ipt", "iam")
    ElseIf FormForScripts.UpdatePart.Value = False And FormForScripts.UpdateAssembly.Value = False And FormForScripts.UpdateDraft.Value = True Then
        ' Только чертежи
        exts = Array("idw", "dwg")
    ElseIf FormForScripts.UpdatePart.Value = False And FormForScripts.UpdateAssembly.Value = True And FormForScripts.UpdateDraft.Value = False = True Then
        ' Только сборки
        exts = Array("iam")
    ElseIf FormForScripts.UpdatePart.Value = True And FormForScripts.UpdateAssembly.Value = False And FormForScripts.UpdateDraft.Value = False = True Then
     ' Только детали
        exts = Array("ipt")
    ElseIf FormForScripts.UpdatePart.Value = True And FormForScripts.UpdateAssembly.Value = False And FormForScripts.UpdateDraft.Value = True = True Then
     ' Только детали и чертежи
        exts = Array("ipt", "idw", "dwg")
    ElseIf FormForScripts.UpdatePart.Value = False And FormForScripts.UpdateAssembly.Value = True And FormForScripts.UpdateDraft.Value = True = True Then
     ' Только сборки и чертежи
        exts = Array("iam", "idw", "dwg")
    Else
        ' По умолчанию — все типы файлов
        exts = Array("idw", "dwg", "ipt", "iam")
    End If

    ProcessFiles folderPath, recursive, exts, rulePath, invApp, iLogicAuto
ThisApplication.SilentOperation = False
    MsgBox "Свойства файлов обновлены!"
End Sub
Private Sub ConvertAllExcelsToPDF()
    Dim excelApp As Object
    Dim wb As Object
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim folderPath As String
    Dim convertedCount As Integer

    folderPath = FormForScripts.PathTB.Text
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Папка не найдена!", vbCritical
        Exit Sub
    End If
    Set folder = fso.GetFolder(folderPath)

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    convertedCount = 0

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or LCase(fso.GetExtensionName(file.Name)) = "xls" Then
            On Error GoTo NextFile
            Set wb = excelApp.Workbooks.Open(file.Path, ReadOnly:=True)
            If Not wb Is Nothing Then
                wb.ExportAsFixedFormat Type:=0, fileName:=Replace(file.Path, ".xlsx", ".pdf")
                wb.Close False
                convertedCount = convertedCount + 1
            End If
NextFile:
            Set wb = Nothing
            Err.Clear
        End If
    Next

    excelApp.Quit
    Set excelApp = Nothing

    MsgBox "Готово! Успешно конвертировано файлов: " & convertedCount, vbInformation
End Sub
Private Sub CreateExcelSpecification()
ThisApplication.SilentOperation = True
    Dim folderPath As String
    folderPath = FormForScripts.PathTB.Text
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim rulePath As String
    rulePath = "C:\\Users\\project15\\Documents\\ilogicRules\\ExcelSpecificationExport.iLogicVb"

    Dim recursive As Boolean
    recursive = FormForScripts.OBallfolders.Value  ' True если выбрана опция "все папки"

    Dim invApp As Application
    Set invApp = GetObject(, "Inventor.Application")

    Dim iLogicAuto As Object
    Set iLogicAuto = GetiLogicAddin(invApp)

    Dim exts As Variant
    exts = Array("iam")

    ProcessFiles folderPath, recursive, exts, rulePath, invApp, iLogicAuto
ThisApplication.SilentOperation = False
    MsgBox "Specification created!"
End Sub


Private Sub ReSaveFile()
ThisApplication.SilentOperation = True
    Dim folderPath As String
    folderPath = FormForScripts.PathTB.Text
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
     Dim rulePath As String
    rulePath = "C:\\Users\\project15\\Documents\\ilogicRules\\ReSave.iLogicVb"
    
    Dim recursive As Boolean
    recursive = FormForScripts.OBallfolders.Value  ' True если выбрана опция "все папки"

    Dim invApp As Application
    Set invApp = GetObject(, "Inventor.Application")

    Dim iLogicAuto As Object
    Set iLogicAuto = GetiLogicAddin(invApp)

    Dim exts As Variant
    exts = Array("idw", "dwg", "ipt", "iam")

    ProcessFiles folderPath, recursive, exts, rulePath, invApp, iLogicAuto
ThisApplication.SilentOperation = False
    MsgBox "Файлы пересохранены!"
End Sub

Public Function GetiLogicAddin(oApplication As Inventor.Application) As Object
    Dim addin As ApplicationAddIn
    On Error GoTo NotFound
    Set addin = oApplication.ApplicationAddIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
    If (addin Is Nothing) Then Exit Function
    addin.Activate
    Set GetiLogicAddin = addin.Automation
    Exit Function
NotFound:
    MsgBox "Не удалось найти iLogic Add-In.", vbCritical
End Function


  Private Sub ProcessFiles(folderPath As String, recursive As Boolean, fileExtensions As Variant, rulePath As String, invApp As Application, iLogicAuto As Object)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fileCount As Integer
    fileCount = 0

    If recursive Then
        ProcessFolderRecursively fso.GetFolder(folderPath), fileExtensions, rulePath, invApp, iLogicAuto, fileCount
    Else
        ProcessSingleFolder folderPath, fileExtensions, rulePath, invApp, iLogicAuto, fileCount
    End If

    MsgBox "Обработка завершена! Всего обработано файлов: " & fileCount, vbInformation
End Sub


Private Sub ProcessSingleFolder(folderPath As String, fileExtensions As Variant, rulePath As String, invApp As Application, iLogicAuto As Object, ByRef fileCount As Integer)
    Dim ext As Variant, fileName As String, doc As Document
    Dim fullFilePath As String
    
    For Each ext In fileExtensions
        fileName = Dir(folderPath & "*." & ext)
        Do While fileName <> ""
            fullFilePath = folderPath & fileName
            Set doc = invApp.Documents.Open(fullFilePath, True)
            If Not doc Is Nothing Then
                iLogicAuto.RunExternalRule doc, rulePath
                doc.Close True
                fileCount = fileCount + 1
            End If
            fileName = Dir
        Loop
    Next ext
End Sub

Private Sub ProcessFolderRecursively(folder As Object, fileExtensions As Variant, rulePath As String, invApp As Application, iLogicAuto As Object, ByRef fileCount As Integer)
    Dim file As Object, subfolder As Object, doc As Document, ext As Variant
    
    ' Пропустить папки с именем  "OldVersions"
    If LCase(folder.Name) = "oldversions" Then Exit Sub

    ' Обработка всех файлов в текущей папке
    For Each file In folder.Files
        For Each ext In fileExtensions
            If LCase(Right(file.Name, Len(ext) + 1)) = "." & ext Then
                Set doc = invApp.Documents.Open(file.Path, True)
                If Not doc Is Nothing Then
                    iLogicAuto.RunExternalRule doc, rulePath
                    doc.Close True
                    fileCount = fileCount + 1
                End If
            End If
        Next ext
    Next file

    ' Рекурсия по папкам
    For Each subfolder In folder.Subfolders
        ProcessFolderRecursively subfolder, fileExtensions, rulePath, invApp, iLogicAuto, fileCount
    Next subfolder
End Sub

