Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Call your macro when the workbook is being saved
    Application.Run "'Save To Github.xlam'!SaveToGithub.AppSaveVbaScriptToGitHub"
End Sub
