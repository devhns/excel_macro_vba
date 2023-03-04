Attribute VB_Name = "Module1"
Sub Word()

Dim wapp As Word.Application
Dim wdoc As Word.Document

Set wapp = CreateObject("word.application")
wapp.Visible = True

Set wdoc = wapp.Documents.Add
wdoc.Content.InsertAfter "MACROCODING"
End Sub
