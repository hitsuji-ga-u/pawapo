
' Clip Path >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub ClipPath()
 Dim MyData As DataObject
 Set MyData = New DataObject
 
 MyData.SetText ActivePresentation.FullName
 MyData.PutInClipboard

End Sub