VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Custom message box"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4695
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButtons 
      Caption         =   "Three buttons"
      Height          =   735
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   2940
      Width           =   1395
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "Two buttons"
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   1620
      Width           =   1395
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "One button"
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1395
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "Custom"
      Height          =   435
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   3240
      Width           =   1275
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "Information"
      Height          =   435
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   1275
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "Exclamation"
      Height          =   435
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "Question"
      Height          =   435
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   1275
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "Error"
      Height          =   435
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "No icon"
      Height          =   435
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   4680
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   4680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButtons_Click(Index As Integer)
    Select Case Index
        Case 0
            MesgBox "Message box with one button.", 5, _
            "Message box title", "Ok"
        Case 1
            MesgBox "Message Box with two buttons" + vbCr + _
            "and two text lines.", 6, "Message box title", _
            "Retry", "Ignore"
        Case 2
            MesgBox "This is a Message Box" + vbCrLf + _
            "with three buttons" + vbCr + "and three text lines", _
            5, "Message box title", "Yes", "No", "Cancel"
            Select Case Message
                Case 0: MsgBox "Button 'Yes' pressed"
                Case 1: MsgBox "Button 'No' pressed"
                Case 2: MsgBox "Button 'Cancel' pressed"
                Case Else
            End Select
        Case Else
    End Select
End Sub

Private Sub cmdIcon_Click(Index As Integer)
    Select Case Index
        Case 0: MesgBox "No icon, no sound", 0, "Message box title", "Ok"
        Case 1: MesgBox "Error", 1, "", "Fix it"
        Case 2: MesgBox "Question", 2, "Message box title", "Do it"
        Case 3: MesgBox "Exclamation", 3, "Message box title", "Good"
        Case 4: MesgBox "Information", 4, "Message box title", "Yes"
        Case 5: MesgBox "Custom icon and sound", 5, "Message box title", "Go"
        Case Else
    End Select
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub
