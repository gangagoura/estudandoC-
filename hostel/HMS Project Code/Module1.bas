Attribute VB_Name = "Module1"
Global cn As New ADODB.Connection

Global sSQL As String
Global MatricNo As String

Sub main()
    'load and display splash screen
    With frmSplash
        .Show
        .Refresh
    End With
    
    'create connection to database to be used throughout the program
    With cn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase\hostel.mdb;Persist Security Info=False"
        .Open
    End With
    
    'display log in form
    frmLogin.Show
    
End Sub
