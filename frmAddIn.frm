VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comment Sync"
   ClientHeight    =   4275
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkResetAllDescriptions 
      Caption         =   "Reset All Descriptions"
      Height          =   372
      Left            =   4260
      TabIndex        =   8
      Top             =   1260
      Width           =   1395
   End
   Begin VB.TextBox txtLog 
      Height          =   2235
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1860
      Width           =   5355
   End
   Begin VB.Frame Frame2 
      Caption         =   "Specific Comment"
      Height          =   675
      Left            =   180
      TabIndex        =   3
      Top             =   1020
      Width           =   3915
      Begin VB.ComboBox comboSign 
         Height          =   300
         Left            =   1500
         TabIndex        =   9
         Text            =   "'''"
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Comment Sign:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apply To"
      Height          =   615
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   3915
      Begin VB.OptionButton OptionProjectGroup 
         Caption         =   "Project Group"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1515
      End
      Begin VB.OptionButton OptionCurrentProject 
         Caption         =   "Current Project"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   4260
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Private mlngCount As Long

Option Explicit

Private Sub CancelButton_Click()
   Connect.Hide
End Sub

Private Sub Form_Load()
   LoadComboSetting comboSign, "'''"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveComboSetting comboSign
End Sub

Private Sub OKButton_Click()
   On Error GoTo hErr
   txtLog.Text = ""
   mlngCount = 0
   
   AddTextToComboList comboSign
   
   If OptionCurrentProject.Value Then
      SyncProjectComments VBInstance.ActiveVBProject
   ElseIf OptionProjectGroup.Value Then
      Dim vbp As VBProject
      For Each vbp In VBInstance.VBProjects
         SyncProjectComments vbp
      Next
   End If
   MsgBox "Updated " & mlngCount & " decription(s).", vbInformation
   Exit Sub
hErr:
   Select Case MsgBox(Err.Description, vbAbortRetryIgnore + vbCritical)
   Case vbRetry
      Resume
   Case vbIgnore
      Resume Next
   End Select
End Sub

Private Sub SyncProjectComments(ByVal vbp As VBProject)
   If vbp Is Nothing Then
      Exit Sub
   End If
   
   Dim comp As VBComponent
   For Each comp In vbp.VBComponents
      If IsAvailableComponent(comp) Then
         ParseComponentComments comp
         
         Dim m As Member
         For Each m In comp.CodeModule.Members
            ParseMemberComments m, comp.CodeModule
         Next
      End If
   Next
End Sub

Private Function IsAvailableComponent(ByVal comp As VBComponent) As Boolean
   Select Case comp.Type
   Case vbext_ct_RelatedDocument, vbext_ct_ResFile
      IsAvailableComponent = False
   Case Else
      If comp.CodeModule Is Nothing Then
         IsAvailableComponent = False
      Else
         If comp.CodeModule.CountOfLines = 0 Then
            IsAvailableComponent = False
         Else
            IsAvailableComponent = True
         End If
      End If
   End Select
End Function

Private Sub ParseComponentComments(ByVal comp As VBComponent)
   Dim cm As CodeModule
   Set cm = comp.CodeModule
   
   If cm Is Nothing Then
      Exit Sub
   End If
   
   Dim i As Long
   i = 1
   
   Dim strComment As String
   strComment = ""
   
   Do While True
      Dim strLine As String
      strLine = LTrim(cm.Lines(i, 1))
      
      If IsCommentLine(strLine) Then
         If GetComment(strLine) <> "" Then
            strComment = strComment & " " & GetComment(strLine)
         End If
      Else
         Exit Do
      End If
      
      i = i + 1
   Loop
   
   If chkResetAllDescriptions.Value Then
      comp.Description = ""
   End If
   
   If strComment <> "" Then
      comp.Description = Trim(strComment)
      
      LogMsg comp.Name
      'LogMsg "[" & comp.Description & "]"
      
      mlngCount = mlngCount + 1
   End If
End Sub

Private Function GetComment(ByVal strLine As String) As String
   GetComment = Trim(Mid(strLine, Len(comboSign.Text) + 1))
End Function

Private Function IsCommentLine(ByVal strLine As String) As Boolean
   Dim lngSignLength As Long
   lngSignLength = Len(comboSign.Text)
   
   If Len(strLine) >= lngSignLength Then
      If LCase(Left(strLine, lngSignLength)) = LCase(comboSign.Text) Then
         IsCommentLine = True
      Else
         IsCommentLine = False
      End If
   Else
      IsCommentLine = False
   End If
End Function

Private Sub ParseMemberComments(ByVal m As Member, ByVal cm As CodeModule)
   On Error GoTo hErr
   
   Dim i As Long
   i = -1
   
   Select Case m.Type
   Case vbext_mt_Method
      i = cm.ProcBodyLine(m.Name, vbext_pk_Proc)
   Case vbext_mt_Property
      If i = -1 Then
         i = cm.ProcBodyLine(m.Name, vbext_pk_Get)
      End If
      
      If i = -1 Then
         i = cm.ProcBodyLine(m.Name, vbext_pk_Set)
      End If
      
      If i = -1 Then
         i = cm.ProcBodyLine(m.Name, vbext_pk_Let)
      End If
   Case Else
      i = m.CodeLocation
   End Select
   
   Debug.Print "Module:" & cm.Parent.Name, "Member:" & m.Name, "Type:" & m.Type, "Loc:" & m.CodeLocation
   
   Dim strComment As String
   strComment = ""
   
   Do While True
      If i <= 1 Then
         Exit Do
      End If
      
      i = i - 1
      
      Dim strLine As String
      strLine = LTrim(cm.Lines(i, 1))
      
      If IsCommentLine(strLine) Then
         If GetComment(strLine) <> "" Then
            strComment = GetComment(strLine) & " " & strComment
         End If
      Else
         Exit Do
      End If
   Loop
   
   Select Case m.Type ' 事件和方法可以重名...
   Case vbext_mt_Event, vbext_mt_Variable, vbext_mt_Method
      strLine = cm.Lines(m.CodeLocation, 1)
      
      Dim pos As Long
      pos = 1
      
      Do While True
         pos = InStr(pos, strLine, comboSign.Text, vbTextCompare)
         
         If pos < 1 Then
            Exit Do
         ElseIf Not IsPosInQuote(strLine, pos) Then
            strLine = Mid(strLine, pos)
            If GetComment(strLine) <> "" Then
               strComment = strComment & " " & GetComment(strLine)
               Exit Do
            End If
         End If
         
         pos = pos + 1
      Loop
   End Select
   
   If chkResetAllDescriptions.Value Then
      m.Description = ""
   End If
   
   If strComment <> "" Then
      m.Description = Trim(strComment)
      
      LogMsg cm.Parent.Name & "." & m.Name
      'LogMsg "[" & m.Description & "]"
   
      mlngCount = mlngCount + 1
   End If
   
   Exit Sub
hErr:
   Select Case Err.Number
   Case 35
   Case Else
      LogMsg "Err " & Err.Number & " " & Err.Description
   End Select
   
   Resume Next
End Sub

Private Sub LogMsg(ByVal strMsg As String)
   txtLog.SelStart = Len(txtLog.Text)
   txtLog.SelText = strMsg & vbCrLf
   txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Function IsPosInQuote(ByVal strExpress As String, ByVal lngPos As Long) As Boolean
   IsPosInQuote = False
   
   If lngPos > Len(strExpress) Then
      Exit Function
   End If
   
   Dim i As Long
   For i = 1 To lngPos
      Dim ch As String
      ch = Mid(strExpress, i, 1)
      
      If ch = """" Then
         IsPosInQuote = Not IsPosInQuote
      End If
   Next
End Function


