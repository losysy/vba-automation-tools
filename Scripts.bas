Attribute VB_Name = "Scripts"

Private Sub ListAllCommands()
    Dim cmdMgr As CommandManager
    Set cmdMgr = ThisApplication.CommandManager

    Dim cmdDef As ControlDefinition
    Dim output As String
    output = ""

    For Each cmdDef In cmdMgr.ControlDefinitions
        output = output & cmdDef.InternalName & " - " & cmdDef.DisplayName & vbCrLf
    Next

    ' Сохраняем результат в текстовый файл
    Dim fso As Object
    Dim txtFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile("C:\Temp\InventorCommandsList.txt", True)
    txtFile.Write output
    txtFile.Close

    MsgBox "Команды сохранены в C:\Temp\InventorCommandsList.txt", vbInformation
End Sub
Private Sub OpenSpec_1()
    ThisApplication.CommandManager.ControlDefinitions.Item("MBD_ADDIN_ToleranceAdvisorViewCmd").Execute
End Sub

Sub ScriptsForm()
 FormForScripts.Show
End Sub


