Imports Autodesk.Navisworks.Api.Automation

Private Sub ConvertToNWD(sender As Object, e As RoutedEventArgs)
    Dim OpenFileDialog1 As New Microsoft.Win32.OpenFileDialog
    OpenFileDialog1.Title = "Select AutoCAD Drawing Files..."
    OpenFileDialog1.Multiselect = True
    OpenFileDialog1.Filter = "Autodesk 2D DWG|*.dwg"
    OpenFileDialog1.ShowDialog()

    If OpenFileDialog1.FileNames.Count > 0 Then
        Dim automationApplication As NavisworksApplication = Nothing
        Dim FileArray As FileInfo() = OpenFileDialog1.FileNames.[Select](Function(f) New FileInfo(f)).ToArray()
        Try
            For Each File In FileArray
                If automationApplication Is Nothing Then
                    automationApplication = New NavisworksApplication()
                    automationApplication.DisableProgress()
                End If
                automationApplication.OpenFile(File.FullName)
                automationApplication.SaveFile(IO.Path.Combine("D:\Temp", IO.Path.GetFileNameWithoutExtension(File.Name) & ".nwf"))
            Next
            If automationApplication IsNot Nothing Then
                automationApplication.EnableProgress()
            End If
        Catch ex As AutomationException
            MessageBox.Show("Error: " & ex.Message)
        Catch ex As AutomationDocumentFileException
            MessageBox.Show("Error: " & ex.Message)
        Finally
            If automationApplication IsNot Nothing Then
                automationApplication.Dispose()
                automationApplication = Nothing
            End If
        End Try

    End If
End Sub
