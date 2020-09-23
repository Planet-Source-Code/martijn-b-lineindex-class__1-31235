VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "LineIndex Demo"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Select 100 Lines"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RANDOM Read 10000 Lines by Number"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private myFile          As New cFileLines
Private lngLinesRead    As Long

Private Sub Command1_Click()
  Dim lngLocation As Long

  On Error Resume Next
  lngLocation = InputBox("Enter a position")
  Text1.Text = myFile.Lines2String(lngLocation, 100)

End Sub

Private Sub Command2_Click()
  Dim sinS             As Single
  Dim sinE             As Single
  Dim lngCount         As Long
  Dim strDummy         As String
  Dim lngMax           As Long
  Dim lngRND           As Long
  
  Me.Cls
  lngMax = 10000
  Randomize Timer
  sinS = Timer
  With myFile
    For lngCount = 1 To lngMax
      lngRND = CLng(Rnd(1) * lngLinesRead) + 1
      strDummy = .Lines2String(lngRND)
    Next
  End With
  sinE = Timer

  Me.Print "#Lines read RANDOMLY:" & lngMax & _
    vbNewLine & "Time: " & sinE - sinS & " s."

End Sub

Private Sub MakeTestFile()
  Dim lngCount        As Long
  Dim strTest As String
  
  On Error Resume Next
  '***  omitted for debugging:
  'Randomize Timer
  Kill App.Path & "\data1.dat"
  
  strTest = "Just some text to test a little...Just some text to test a little...Just some text to test a little..."
  Open App.Path & "\data1.dat" For Output As #1
  For lngCount = 1 To 100000
    Print #1, "This is line# " & Format(lngCount, "00000") & " >>" & Left$(strTest, CLng(Rnd(1) * 70) + 1) & "<<"
  Next
  Close

End Sub

Private Sub Form_Activate()
  Dim sinS             As Single
  Dim sinE             As Single

  With myFile
    Me.Print "Scanning file...": DoEvents
    sinS = Timer
    lngLinesRead = .ReadFile(App.Path & "\data1.dat")
    sinE = Timer
    Me.Print "File Size: " & myFile.LengthOfFile & " bytes"
    Me.Print lngLinesRead & " lines in file.": DoEvents
    Me.Print "Memory consumption: " & lngLinesRead * 4 + 2000 & " bytes": DoEvents
    Me.Print "Time: " & sinE - sinS & " s.": DoEvents
  End With
End Sub

Private Sub Form_Load()
  MakeTestFile
End Sub

Private Sub Form_Terminate()
  On Error Resume Next
  Set myFile = Nothing
  Kill App.Path & "\data1.dat"
End Sub
