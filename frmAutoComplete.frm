VERSION 5.00
Begin VB.Form frmAutoComplete 
   Caption         =   "Auto Complete Example"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cboAuto 
      Height          =   315
      ItemData        =   "frmAutoComplete.frx":0000
      Left            =   120
      List            =   "frmAutoComplete.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmAutoComplete.frx":0004
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnAuto As Boolean 'Keeps the autocomplete functions from
                        'triggering the Change event

Private Sub cboAuto_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    'don't do if no text or if change was made by autocomplete coding
    If Not blnAuto And cboAuto.Text <> "" Then
        'save the selection start point (cursor position)
        iStart = cboAuto.SelStart
        'get the part the user has typed (not selected)
        strPart = Left$(cboAuto.Text, iStart)
        For iLoop = 0 To cboAuto.ListCount - 1
            'compare each item to the part the user has typed,
            '"complete" with the first good match
            strItem = UCase$(cboAuto.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(cboAuto.Text) Then
                'partial match but not the whole thing.
                '(if whole thing, nothing to complete!)
                blnAuto = True
                cboAuto.SelText = Mid$(cboAuto.List(iLoop), iStart + 1) 'add on the new ending
                cboAuto.SelStart = iStart   'reset the selection
                cboAuto.SelLength = Len(cboAuto.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub cboAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    'Unless we watch out for it, backspace or delete will just delete
    'the selected text (the autocomplete part), so we delete it here
    'first so it doesn't interfere with what the user expects
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        cboAuto.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        cboAuto_LostFocus
        'the following causes the item to be selected and
        'the cursor placed at the end:
        cboAuto.SelStart = Len(cboAuto.Text)
        
        'This would select the whole thing instead:
        'cboAuto.SelLength = Len(cboAuto.Text)
    
        'alternatively, you could move the focus to the next control here
    End If
End Sub

Private Sub cboAuto_LostFocus()
Dim iLoop As Integer
'Match capitalization if item entered is one on the list
    If cboAuto.Text <> "" Then
        For iLoop = 0 To cboAuto.ListCount - 1
            If UCase$(cboAuto.List(iLoop)) = UCase$(cboAuto.Text) Then
                blnAuto = True
                cboAuto.Text = cboAuto.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Form_Load()
    'add a bunch of items.  cboAuto's Sorted property is
    'True so they will end up in order
    cboAuto.AddItem "Apples"
    cboAuto.AddItem "Oranges"
    cboAuto.AddItem "Bananas"
    cboAuto.AddItem "Pears"
    cboAuto.AddItem "Peaches"
    cboAuto.AddItem "Pineapples"
    cboAuto.AddItem "Grapes"
    cboAuto.AddItem "Blueberries"
    cboAuto.AddItem "Raspberries"
    cboAuto.AddItem "Blackberries"
    cboAuto.AddItem "Papaya"
    cboAuto.AddItem "Kiwi"
    cboAuto.AddItem "Watermelon"
    cboAuto.AddItem "Guava"
End Sub
