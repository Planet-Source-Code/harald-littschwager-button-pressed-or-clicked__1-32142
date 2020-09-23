VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "&Mouse"
      Height          =   510
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   720
      Width           =   4380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Press Enter"
      Height          =   510
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   4380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'checks if user presed ENTER on a CommandButton or used the MOUSE
'ALT+Key=Mouseclick

Private Sub Command1_Click(Index As Integer)
Dim Dummy As String     'only for the Msgbox

Select Case Index       'which button is pressed
    
    Case 0              'User should presse 'Enter'
        Dummy = IIf(ReturnState, "That's fine", "Don't click, press 'Enter'")
        Command1(1).SetFocus
    Case 1              'User should use the mouse
        Dummy = IIf(ReturnState, "Don't use keyboard, use mouse !", "That's OK")
        Command1(0).SetFocus
End Select
MsgBox Dummy            'say what's happened
End Sub

