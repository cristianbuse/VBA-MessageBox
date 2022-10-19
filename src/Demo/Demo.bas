Attribute VB_Name = "Demo"
Option Explicit

Public Sub DemoMain()
    Debug.Print MessageBox("Test", "Title", icoCritical, "Button1", "Button2")
    Debug.Print MessageBox("Test", , icoInformation, "Merge", "Replace", "Ignore", "Abort", "Choose", 3)
    Debug.Print MessageBox(promptText:="Test" _
                         , titleText:="Some title" _
                         , ico:=icoQuestion _
                         , button1:="This is the first looooooooooooooooooooooooooong question" _
                         , button2:="This is the second loooooooooooooooooooooooooong question" _
                         , button3:="This is the third looooooooooooooooooooooooooong question")
    Debug.Print MessageBox("Test") 'Displays OK only and allows Cancel via X or Esc
    MessageBox.Show "Test", , icoExclamation
    '
    Dim lineText As String
    Dim i As Long
    Dim s As String
    '
    lineText = Join(Array(String(10, "A"), String(15, "B"), String(25, "C")), ", ")
    For i = 1 To 100
        s = s & lineText & vbNewLine
    Next i
    MessageBox.Show s 'Displays vertical scroll bar
    MessageBox.Show String(100, "B") & s 'Displays both vertical and horizontal scroll bars
End Sub
