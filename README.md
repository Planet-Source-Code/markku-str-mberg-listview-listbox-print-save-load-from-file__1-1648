<div align="center">

## ListVieW/ListBox Print, Save, Load from File


</div>

### Description

This code lets you to Print, Save to file, Load data from file to/from ListBox or ListView object.
 
### More Info
 
Listbox related functions:

LoadListBox -Inputs: ListBox object, Directory from where to load

SaveListBox -Inputs: ListBox object, Directory where to save

PrintListBox -Inputs: ListBox object

ListView related:

SaveLV -Inputs: ListView object, Subitems amount in that object, Directory where to save

PrintLV -Inputs: Listview object, Subitems amount in that object

I fogot to say that those ListView functions assumes that ListView viewmode is lvwReport and you should have one or more subitems in that object, Sorry ;)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Markku Strömberg](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/markku-str-mberg.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/markku-str-mberg-listview-listbox-print-save-load-from-file__1-1648/archive/master.zip)





### Source Code

```
'Example: Call SaveListBox(list1, "C:\Temp\MyList.dat")
Public Sub SaveListBox(TheList As ListBox, Directory As String)
 Dim SaveList As Long
 On Error Resume Next
 Open Directory$ For Output As #1
 For SaveList& = 0 To TheList.ListCount - 1
  Print #1, TheList.List(SaveList&)
 Next SaveList&
 Close #1
End Sub
'Example: Call LoadListBox(list1, "C:\Temp\MyList.dat")
Public Sub LoadListBox(TheList As ListBox, Directory As String)
 Dim MyString As String
 On Error Resume Next
 Open Directory$ For Input As #1
 While Not EOF(1)
  Input #1, MyString$
   DoEvents
    TheList.AddItem MyString$
 Wend
 Close #1
End Sub
Public Sub PrintListBox(TheList As ListBox)
 Dim SaveList As Long
 On Error Resume Next
 Printer.FontSize = 12
 For SaveList& = 0 To TheList.ListCount - 1
  Printer.Print TheList.List(SaveList&)
 Next SaveList&
 Printer.EndDoc
End Sub
Public Function PrintLV(lv As ListView, Subs As Integer)
 Printer.FontSize = 12
 Dim subit As Variant
 Dim i As Integer
 Dim x As Integer
 For i = 1 To lv.ListItems.Count
  subit = lv.ListItems(i).Text & vbTab
  For x = 1 To Subs
   subit = subit & lv.ListItems(i).SubItems(x) & vbTab
  Next
  Printer.Print subit
  subit = ""
 Next
 Printer.EndDoc
End Function
Public Function SaveLV(lv As ListView, Subs As Integer, sPath As String)
 Dim subit As Variant
 Dim F As Integer
 Dim i As Integer
 Dim x As Integer
 F = FreeFile
 On Error Resume Next
 Open sPath For Output As #F
 For i = 1 To lv.ListItems.Count
  subit = lv.ListItems(i).Text & vbTab
  For x = 1 To Subs
   subit = subit & lv.ListItems(i).SubItems(x) & vbTab
  Next
  Print #F, subit
  subit = ""
 Next
 Close #F
End Function
```

