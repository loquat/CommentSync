Attribute VB_Name = "GlobalCommon"
Option Explicit

Public fso As New FileSystemObject

Public Sub ClearComboList(ByVal combo As ComboBox)
   Dim sTemp As String
   sTemp = combo.Text
   
   combo.Clear
   
   combo.Text = sTemp
End Sub

Public Function AddTextToComboList(ByVal combo As ComboBox) As Boolean
   If combo.Text = "" Then
      Exit Function
   End If

   Dim bFound As Boolean
   bFound = False
   
   Dim i As Long
   For i = 1 To combo.ListCount
      If LCase(combo.List(i - 1)) = LCase(combo.Text) Then
         bFound = True
         Exit For
      End If
   Next
   
   If Not bFound Then
      combo.AddItem combo.Text, 0
      AddTextToComboList = True
   Else
      AddTextToComboList = False
   End If
End Function

Public Sub SaveComboSetting(ByVal combo As ComboBox, Optional ByVal bTextOnly As Boolean = False)
   With combo
      SaveSetting App.EXEName, "main", .Name & ".Text", .Text
      If Not bTextOnly Then
         SaveSetting App.EXEName, "main", .Name & ".ListCount", .ListCount
         
         Dim i As Long
         For i = 1 To .ListCount
            SaveSetting App.EXEName, "main", .Name & ".ListItem" & i, .List(i - 1)
         Next
      End If
   End With
End Sub

Public Sub LoadComboSetting( _
   ByVal combo As ComboBox, _
   ByVal default As String, _
   Optional ByVal bTextOnly As Boolean = False)
   
   With combo
      If Not bTextOnly Then
         Dim nCount As Long
         nCount = GetSetting(App.EXEName, "main", .Name & ".ListCount", "0")
         
         .Clear
         
         Dim i As Long
         For i = 1 To nCount
            .AddItem GetSetting(App.EXEName, "main", .Name & ".ListItem" & i)
         Next
      End If
   
      .Text = GetSetting(App.EXEName, "main", .Name & ".Text", default)
   End With
End Sub

Public Function IsFormLoaded(ByVal frm As Form) As Boolean
   IsFormLoaded = False
   
   Dim f As Form
   For Each f In Forms
      If f Is frm Then
         IsFormLoaded = True
         Exit Function
      End If
   Next
End Function

Public Sub SaveCheckBoxSetting(ByRef chk As CheckBox)
   SaveSetting App.EXEName, "main", chk.Name, chk.Value
End Sub

Public Sub LoadCheckBoxSetting(ByRef chk As CheckBox, ByVal default As CheckBoxConstants)
   chk.Value = GetSetting(App.EXEName, "main", chk.Name, default)
End Sub

Public Sub SaveListSetting(ByRef lst As ListBox)
   With lst
      SaveSetting App.EXEName, "main", .Name & ".ListCount", .ListCount
      
      Dim i As Long
      For i = 1 To .ListCount
         SaveSetting App.EXEName, "main", .Name & ".ListItem" & i, .List(i - 1)
      Next
      If lst.Style = vbListBoxCheckbox Then
         For i = 1 To .ListCount
            SaveSetting App.EXEName, "main", .Name & ".Selected" & i, .Selected(i - 1)
         Next
      End If
   End With
End Sub

Public Sub LoadListSetting(ByRef lst As ListBox)
   With lst
      Dim nCount As Long
      nCount = GetSetting(App.EXEName, "main", .Name & ".ListCount", "0")
      
      .Clear
      
      Dim i As Long
      For i = 1 To nCount
         .AddItem GetSetting(App.EXEName, "main", .Name & ".ListItem" & i)
         .Selected(.NewIndex) = GetSetting(App.EXEName, "main", .Name & ".Selected" & i)
      Next
      If lst.Style = vbListBoxCheckbox Then
         For i = 1 To nCount
            .Selected(i) = GetSetting(App.EXEName, "main", .Name & ".Selected" & i)
         Next
      End If
   End With
End Sub

Public Function AddStrToList(ByRef lst As ListBox, ByVal newstr As String) As Boolean
   AddStrToList = False
   Dim i As Long
   For i = 1 To lst.ListCount
      If LCase(lst.List(i - 1)) = LCase(newstr) Then
         Exit Function
      End If
   Next
   
   lst.AddItem newstr
   AddStrToList = True
End Function
