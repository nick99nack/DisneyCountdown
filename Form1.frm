VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Countdown to Disney"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DatePicker 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   44733
      MaxDate         =   73050
      MinDate         =   20285
   End
   Begin VB.Label CountdownLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Countdown Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vacationDate As Date
Dim dateDifference As Integer

Private Sub DatePicker_CloseUp()
vacationDate = DatePicker.Value
UpdateCaption

End Sub

Private Sub ExitButton_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    ChgBackground
End Sub

Private Sub Form_Load()
    Dim imgNum As Integer
    ReadFile
    UpdateCaption
    ChgBackground
End Sub

Public Sub UpdateCaption()

    dateDifference = DateDiff("d", Now, vacationDate)
    If dateDifference < 0 Then
        CountdownLabel.Caption = Abs(dateDifference) & " days since Disney"
    ElseIf dateDifference = 1 Then
        CountdownLabel.Caption = Abs(dateDifference) & " days since Disney"
    Else
        CountdownLabel.Caption = dateDifference & " days until Disney!"
    End If

End Sub

Public Sub WriteFile()

    Dim intFile As Integer
    Dim strFile As String
    strFile = App.Path & "\save.dat"
    intFile = FreeFile
    Open strFile For Output As #intFile
        Print #intFile, CLng(vacationDate)
    Close #intFile
    
End Sub

Public Sub ReadFile()

    Dim intFile As Integer
    Dim strFile As String
    Dim strFileContents As String
    strFile = App.Path & "\save.dat"
    intFile = FreeFile
    Open strFile For Input As #intFile
        Input #intFile, strFileContents
    Close #intFile
    If Not strFileContents < 20285 And Not strFileContents > 73050 Then
        DatePicker.Value = CVDate(strFileContents)
        vacationDate = DatePicker.Value
    Else
        DatePicker.Value = Now
        vacationDate = DatePicker.Value
    End If
        
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFile
End Sub

Public Sub ChgBackground()
    Randomize
    imgNum = Int((22 - 1 + 1) * Rnd + 1)
    Form1.Picture = LoadPicture(App.Path & "\img\" & imgNum & ".bmp")
End Sub
