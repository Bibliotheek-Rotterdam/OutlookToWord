Attribute VB_Name = "Belastingdienst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub OpenInWord()
    On Error Resume Next
    Dim MyWindow As Outlook.Inspector
    Dim MyItem As MailItem
    Dim FilePath As String
    FilePath = Environ("TEMP") & "\"
    Dim ItemName As String
    Set MyWindow = Application.ActiveInspector
    
    If TypeName(MyWindow) = "Nothing" Then
        MsgBox ("Open een e-mailtje om deze in Word te openen.")
    Else
        Set MyItem = MyWindow.CurrentItem
        ItemName = MyItem.Subject
        With MyItem
            .SaveAs FilePath & ItemName & ".html", olHTML
        End With
        
        Dim WordApp As Object, WordDoc As Object
        Set WordApp = CreateObject("Word.Application")
        WordApp.Visible = True
        WordDoc.Documents.Add
        Set WordDoc = WordApp.Documents.Open(FilePath & ItemName & ".html")
    End If
End Sub
