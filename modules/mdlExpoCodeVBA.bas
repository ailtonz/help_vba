Attribute VB_Name = "mdlExpoCodeVBA"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
' https://gist.github.com/steve-jansen/7589478
' ------------------------------------------------------------------
' Adaptado e traduzido por Sidnei Graciolli - Dez/2017
' Macro para exportar todo o código VBA deste projeto para arquivos de texto para controle de versionamento
' Requer a habilitação da opção de segurança no Excel no menu Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
' ------------------------------------------------------------------
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    If Not ThisWorkbook.Saved Then
        MsgBox "Salve o trabalho antes de exportar o projeto VBA", vbCritical, "ExpoCodeVBA"
        Exit Sub
    End If
    
    directory = ThisWorkbook.path & "\" & Split(ThisWorkbook.Name, ".")(0) & "_VBA"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ThisWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Erro ao exportar o componente " & VBComponent.Name & " para " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exportado " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    MsgBox "Exportação com sucesso de " & CStr(count) & " arquivos VBA para " & directory, vbInformation, "ExpoCodeVBA"
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub
