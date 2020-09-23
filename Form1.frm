VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dynamic Popup Menu"
   ClientHeight    =   2745
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add Item"
      Height          =   315
      Left            =   1830
      TabIndex        =   3
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "Menu Item 1"
      Top             =   60
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   60
      TabIndex        =   1
      Top             =   390
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1830
      TabIndex        =   4
      Top             =   450
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Right Click on the Form"
      Height          =   195
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Menu myMenu 
      Caption         =   ""
      Begin VB.Menu myMenuArray 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Name : Dynamic Popup Menu
'Description : Dynamically Adds Items to PopUp Menu
'By :  Leartes  -  leartes@leartes.net
'
'Prepration :
'Step 1: Add Menu with Menu Editor to the Form
'Step 2: Add an Item to Menu without Caption (in Example Named myMenu)
'Step 3: Add an Sub Item to Menu without Caption and to Enable Array give Index value 0 (in Example Named myMenuArray)
'
'I Hope it gives an idea, how to add items dynamically to a popup menu
'***********************************************************************************

Private Sub Command1_Click()
If Text1.Text <> "" Then List1.AddItem Text1.Text: Label1.Visible = True: Text1.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Call AddtoDynPop
    If myMenuArray.UBound > 0 Then
    myMenuArray(myMenuArray.UBound).Visible = False
    Me.PopupMenu myMenu
    End If
End If
End Sub
Private Sub AddtoDynPop()
For I = myMenuArray.LBound + 1 To myMenuArray.UBound
Unload myMenuArray(I)
Next
For I = 0 To List1.ListCount - 1
Load myMenuArray(I + 1)
myMenuArray(I).Caption = List1.List(I)
Next

End Sub

Private Sub myMenuArray_Click(Index As Integer)
Label2.Caption = """" & myMenuArray(Index).Caption & """ Selected"
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
