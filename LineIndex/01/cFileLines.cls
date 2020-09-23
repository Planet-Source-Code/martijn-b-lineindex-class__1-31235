VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private parrLineBuffer() As Long '*** Line end location array
Private plngBufferLen   As Long '*** Ubound of parrLineBuffer()
Private plngLinesRead   As Long '*** # lines in parrLineBuffer()
Private pblnInitialized As Boolean '***  True if ReadFile succeeded
Private plngFile        As Long '***  Stores the locked file number

Public Function ReadFile(strFileName As String) As Long
  Dim lngBlockSize     As Long '***  Length of file read buffer
  Dim lngLOF           As Long '***  Length of file
  Dim lngFilePos       As Long '***  Start Pos of last block read
  Dim lngNoBlocks      As Long '***  lngNoBlocks = lngLOF \ lngBlockSize
  Dim lngExtra         As Long '***  lngLOF Mod lngBlockSize
  Dim lngMain          As Long '***  lngMain = lngNoBlocks * lngBlockSize
  Dim arrBlock()       As Byte '***  file read buffer
  Dim lngChar          As Long '***  search position in file read buffer
  Dim lngBlockLen      As Long '***  lngBlockLen = UBound(arrBlock) + 1&
  Dim lngCurrentState  As Long '***  1 if last byte of arrBlock = vbLF
  Dim arrDelimit(1)    As Byte '***  stores vbNewLine for InstrB use

  '***  If called a second time, cleanup
  pblnInitialized = False
  Erase parrLineBuffer
  plngLinesRead = 0&
  plngBufferLen = 0&
  Close plngFile
  plngFile = 0&

  '***  put the end of line in the byte array
  arrDelimit(0) = 13
  arrDelimit(1) = 10

  '***  initialize the buffer; there maybe just one line
  ReDim parrLineBuffer(plngLinesRead)

  '***  amount of memory used for file read, independend of the file size
  lngBlockSize = 2000&
  '***  instr base compatibility: the ' -1& ' in code would slow down the parser
  ReDim arrBlock(1& To lngBlockSize)

  plngFile = FreeFile
  '***  use the whole vocabulary :)
  Open strFileName For Binary Access Read Lock Write As #plngFile
  lngLOF = LOF(plngFile)

  lngNoBlocks = lngLOF \ lngBlockSize
  lngExtra = lngLOF Mod lngBlockSize
  lngMain = lngNoBlocks * lngBlockSize

  lngCurrentState = 0&

  '***  do not load the entire file here
  '***  cause that could cause memory issues
  For lngFilePos = 1& To lngMain Step lngBlockSize
    Get #plngFile, lngFilePos, arrBlock()
    '***  sorry but gosub is faster than a function call
    GoSub tagSearch
  Next

  '***  read the last bit of the file
  If lngExtra <> 0 Then
    ReDim arrBlock(1& To lngExtra)
    Get #plngFile, lngFilePos, arrBlock()
    '***  sorry but gosub is faster than a function call
    GoSub tagSearch
  End If

  '***  store the EOF as a line end if last line has no vbNewLine
  If lngLOF - 1 <> parrLineBuffer(plngLinesRead) Then
    GoSub tagStoreLine
    '***  correct for missing vbNewLine
    parrLineBuffer(plngLinesRead) = parrLineBuffer(plngLinesRead) - 2
  End If
  '***  store the BOF as a line end of 'virtual' line #0
  parrLineBuffer(0) = -1
  '***  remove extra buffer space
  ReDim Preserve parrLineBuffer(0 To plngLinesRead)

  ReadFile = plngLinesRead

  If lngLOF > 0 Then
    '***  allow reading of lines
    pblnInitialized = True
  End If

  Exit Function

tagSearch:
  lngChar = 1&

  '***  check for broken vbNewline
  If lngCurrentState = 1& Then
    If arrBlock(1&) = arrDelimit(1&) Then
      '***  found a delimiter
      'Debug.Print "found at: " & lngFilePos + lngChar - 1&
      lngChar = 0&
      GoSub tagStoreLine
      lngCurrentState = 0&
      lngChar = 2&
    End If
  End If

  lngBlockLen = UBound(arrBlock) + 1&

  Do

    '***  searching
    lngChar = InStrB(lngChar, arrBlock, arrDelimit, vbBinaryCompare)

    If lngChar = 0& Then
      lngChar = lngBlockLen
    Else
      '***  found a delimiter
      'Debug.Print "found at: " & lngFilePos + lngChar - 1&
      GoSub tagStoreLine
      lngChar = lngChar + 1&
    End If

  Loop Until lngChar = lngBlockLen

  '***  check for broken vbNewline
  If arrBlock(lngBlockLen - 1&) = arrDelimit(0) Then
    '***  delimit was cut
    lngCurrentState = 1&
  End If
  Return
  '***  End tagSearch:

tagStoreLine:
  plngLinesRead = plngLinesRead + 1&
  '***  resize the buffer if needed
  If plngBufferLen < plngLinesRead Then
    plngBufferLen = plngBufferLen * 2 + 100
    ReDim Preserve parrLineBuffer(0 To plngBufferLen) As Long
  End If
  '***  store the end of the line loc
  parrLineBuffer(plngLinesRead) = lngFilePos + lngChar - 1&
  Return
  '***  End tagStoreLine:

End Function

Public Function Lines2String(lngLineNumber As Long, Optional lngNumber As Long = 1) As String
  Dim arrTempLineBuffer() As Byte
  Dim lngUbound        As Long '***  UpperBound of buffer
  Dim lngStart         As Long '***  Start of read
  Dim lngEnd           As Long '***  End of read
  If pblnInitialized Then
    '*** check input
    If (lngLineNumber > 0) And ((lngLineNumber + lngNumber - 1) <= plngLinesRead) Then

      lngStart = parrLineBuffer(lngLineNumber - 1&) + 2
      lngEnd = parrLineBuffer(lngLineNumber + lngNumber - 1) + 1

      '***  read the lines
      ReDim arrTempLineBuffer(lngEnd - lngStart) As Byte
      Get #plngFile, lngStart, arrTempLineBuffer()

      '***  remove the last vbNewLine for 'Line Input #1, strLine' - compatibility
      '***  the class is 20% faster if you remove this
      lngUbound = UBound(arrTempLineBuffer)
      If lngUbound < 2& Then
        '*** just a vbNewLine; return empty string
        Lines2String = vbNullString
      ElseIf (lngLineNumber + lngNumber - 1) = plngLinesRead Then
        '***  return a string
        Lines2String = StrConv(arrTempLineBuffer, vbUnicode)
      Else
        '***  remove vbNewLine
        ReDim Preserve arrTempLineBuffer(lngUbound - 2)

        '***  return a string
        Lines2String = StrConv(arrTempLineBuffer, vbUnicode)

      End If
    End If
  End If
End Function

Public Function Lines2ByteArray(lngLineNumber As Long, Optional lngNumber As Long = 1) As Byte()
  '***  Returns the line(s) without any modification
  Dim lngStart         As Long '***  Start of read
  Dim lngEnd           As Long '***  End of read
  Dim arrTempLineBuffer() As Byte
  
  If pblnInitialized Then
    '*** check input
    If (lngLineNumber > 0) And ((lngLineNumber + lngNumber - 1) <= plngLinesRead) Then

      lngStart = parrLineBuffer(lngLineNumber - 1&) + 2
      lngEnd = parrLineBuffer(lngLineNumber + lngNumber - 1) + 1

      '***  read the lines
      ReDim arrTempLineBuffer(lngEnd - lngStart) As Byte
      Get #plngFile, lngStart, arrTempLineBuffer()

      '***  return a byte array
      Lines2ByteArray = arrTempLineBuffer
      Erase arrTempLineBuffer

    End If
  End If
End Function

Private Sub Class_Terminate()
  If pblnInitialized Then
    Erase parrLineBuffer
    Close plngFile
  End If
End Sub

Public Property Get LengthOfFile() As Long
  If pblnInitialized Then
    LengthOfFile = LOF(plngFile)
  End If
End Property
