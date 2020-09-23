Attribute VB_Name = "modMessgBox"
Option Explicit

'******************************************************
'
' MessgBox module
' Author: Raul Lopez  raulopez@hotmail.com
' Use this code as you like. It isn't necessary
' to mention me nor give me any credit. Enjoy it!
'
'******************************************************
'
'You can use any button, included standard VB CommdButt
'You can use any caption for your buttons
'You can use standard or custom icons
'You can use system or custom sounds
'
'******************************************************
'
'To call the function:
'
'MesgBox "Text", n, "Title", "Capt1", "Capt2", "Capt3"
'
'Where
'   Text = Message
'      n = Icon and sound value
'  Title = Message box title
'  Capt1 = First button caption
'  Capt2 = Second button caption (optional)
'  Capt3 = Third button caption (optional)
'
'******************************************************
'
'The value for the icons and sounds are:
'  0 = None
'  1 = Critical
'  2 = Question
'  3 = Exlamation
'  4 = Information
'  5 = Custom
'and next
'
'The button's value are:
'  0 = First button
'  1 = Second button
'  2 = Third button
'
'******************************************************

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
       (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Public Const MB_ICONASTERISK = &H10&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONINFORMATION = &H40&

Public IcoVal As Integer
Public Message As Integer

Public Function MesgBox(mbText As String, mbIcon As Integer, _
                        mbTitle As String, mbCBut0 As String, _
                        Optional mbCBut1 As String, _
                        Optional mbCBut2 As String) _
                        As Long

   IcoVal = mbIcon  'Stores icon value to play sound
   
   With frmMessgBox
      'Message
      .Label1.Caption = mbText
      
      'Message box title
      If mbTitle <> "" Then
         .Caption = mbTitle
      Else
         .Caption = App.Title
      End If
      
      'Buttons caption
      .Button(0).Caption = mbCBut0
      .Button(1).Caption = mbCBut1
      .Button(2).Caption = mbCBut2
      
      'Icon
      If mbIcon <> 0 Then
         .Image1.Picture = .ImageList1.ListImages(mbIcon).Picture
      End If
      
      If mbCBut1 <> "" Then .Button(1).Visible = True
      If mbCBut2 <> "" Then .Button(2).Visible = True
      
      'One button
      .Button(0).Left = 2130
      
      'Two buttons
      If mbCBut1 <> "" And mbCBut2 = "" Then
         .Button(0).Left = 1320
         .Button(1).Left = 2940
      End If
    
      'Three buttons
      If mbCBut1 <> "" And mbCBut2 <> "" Then
         .Button(0).Left = 840
         .Button(1).Left = 2130
         .Button(2).Left = 3420
      End If

   End With
   
   frmMessgBox.Show 1

End Function
