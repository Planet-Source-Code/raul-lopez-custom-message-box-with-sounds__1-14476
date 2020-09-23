VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title"
   ClientHeight    =   1485
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   4605
   ControlBox      =   0   'False
   Icon            =   "MessgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin MessgBox.XDWButton Button 
      Height          =   345
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1050
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Caption"
      HoverColor      =   -2147483639
      ButtonBorderStyle=   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MessgBox.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MessgBox.frx":0330
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MessgBox.frx":0654
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MessgBox.frx":0978
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MessgBox.frx":0C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MessgBox.frx":1578
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MessgBox.XDWButton Button 
      Height          =   345
      Index           =   1
      Left            =   2130
      TabIndex        =   1
      Top             =   1050
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Caption"
      HoverColor      =   -2147483639
      ButtonBorderStyle=   1
   End
   Begin MessgBox.XDWButton Button 
      Height          =   345
      Index           =   2
      Left            =   3420
      TabIndex        =   2
      Top             =   1050
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Caption"
      HoverColor      =   -2147483639
      ButtonBorderStyle=   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   4620
      X2              =   720
      Y1              =   940
      Y2              =   940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   4620
      X2              =   720
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Message box text"
      Height          =   675
      Left            =   900
      TabIndex        =   3
      Top             =   180
      Width           =   3555
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   4620
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   300
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      Index           =   1
      X1              =   728
      X2              =   728
      Y1              =   0
      Y2              =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   720
      X2              =   720
      Y1              =   0
      Y2              =   1500
   End
End
Attribute VB_Name = "frmMessgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Button_Click(Index As Integer)
'Stores value of pressed button
   Select Case Index
      Case 0: Message = 0
      Case 1: Message = 1
      Case 2: Message = 2
      Case Else 'Nothing to do
   End Select
   Unload Me
End Sub

Private Sub Form_Load()
'Plays a sound according to selected value
   Dim CustSnd As Long
   Select Case IcoVal
      Case 1: MessageBeep MB_ICONASTERISK
      Case 2: MessageBeep MB_ICONQUESTION
      Case 3: MessageBeep MB_ICONEXCLAMATION
      Case 4: MessageBeep MB_ICONINFORMATION
      Case 5: CustSnd = sndPlaySound(App.Path & "\Tada.wav", 1)
      Case 6: CustSnd = sndPlaySound(App.Path & "\Test.wav", 1)
      Case Else 'Nothing to do
   End Select
End Sub

