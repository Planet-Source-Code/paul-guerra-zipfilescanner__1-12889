Attribute VB_Name = "modZip"
Option Explicit
Const Sign1 = &H4034B50, Sign2 = &H2014B50, Sign3 = &H6054B50, Bit1 = 1, Bit3 = 4
Private Type LocalFileHeader
  Version As Integer
  GeneralFlags As Integer
  Method As Integer
  ModifiedTime As Integer
  ModifiedDate(1) As Byte
  Crc As Long
  CompSize As Long
  UncompSize As Long
  FileLen As Integer
  ExtraLen As Integer
End Type
Private Type CentralDirectory
  Version As Integer
  VersionExtract As Integer
  GeneralFlags As Integer
  Method As Integer
  Modified As Long
  Crc As Long
  CompSize As Long
  UncompSize As Long
  FileLen As Integer
  ExtraLen As Integer
  FileCommentLen As Integer
  DiskStart As Integer
  IntAttr As Integer
  ExtAttr As Long
  RelOffsetLocalHeader As Long
End Type
Private Type EndCentralDirectory
  DiskNum As Integer
  CentralDirDisk As Integer
  CentralDirEntries As Integer
  CentralDirTotalEntries As Integer
  CentralDirSize As Long
  CentralDirStart As Long
  ZipComment As String
End Type
Private Type UseFullInfo
  ExtraField As String
  FileComment As String
  CompSize As Long
  UncompSize As Long
  Crc As Long
  Nothing As Boolean
  'Modification info
  Day As Byte
  Month As Byte
  Year As Integer
  Hour As Byte
  Minute As Byte
  Seconds As Byte
End Type
Private Type DataDescriptor
  Crc As Long
  CompSize As Long
  UncompSize As Long
End Type
Dim Info() As UseFullInfo
Dim ZipComment As String
  
Sub LoadFiles(File As String)
  Dim LocalHeader As LocalFileHeader
  Dim CentralDir As CentralDirectory
  Dim EndCentralDir As EndCentralDirectory
  Dim Descriptor As DataDescriptor
  Dim Sign As Long, i As Long, Parent As Long, LastParent As Long, Counter As Long
  Dim Matrix() As String
  Dim LoadingPhase As Boolean

  ReDim Info(0)
  With frmMain.usrZipFiles
    .Clear
    .DoNotRefresh = True
    Open File For Binary Access Read As #1
    Do
      Get #1, , Sign
      Select Case Sign
        Case Sign1
          If LoadingPhase Then
            MsgBox "Local File Header singnature found in the wrong place", vbCritical
            Close
            Exit Sub
          End If
          Get #1, , LocalHeader
          Separate Input(LocalHeader.FileLen, 1), Matrix()
          If Len(Matrix(UBound(Matrix))) Then
            If LocalHeader.GeneralFlags And Bit1 Then Matrix(UBound(Matrix)) = Matrix(UBound(Matrix)) + "  ?"
            LastParent = -1
            For i = 0 To UBound(Matrix)
              Parent = .SearchItem(Matrix(i), i)
              If Parent = -1 Then
                LastParent = .AddItem(LastParent, Matrix(i))
                ReDim Preserve Info(UBound(Info) + 1)
                Info(UBound(Info)).Nothing = True
              Else
                LastParent = Parent
              End If
            Next i
            DateConv LocalHeader.ModifiedDate(), Info(UBound(Info))
            TimeConv LocalHeader.ModifiedTime, Info(UBound(Info))
            With Info(UBound(Info))
              .Nothing = False
              .CompSize = LocalHeader.CompSize
              .UncompSize = LocalHeader.UncompSize
              .Crc = LocalHeader.Crc
              .ExtraField = Input(LocalHeader.ExtraLen, 1)
            End With
          End If
          Seek #1, Seek(1) + LocalHeader.CompSize
          If LocalHeader.GeneralFlags And Bit3 Then Get #1, , Descriptor
        Case Sign2
          If Counter > UBound(Info) Then
            MsgBox "In the Central Directory there are more file entries than actual files", vbCritical
            Close
            Exit Sub
          End If
          LoadingPhase = True
          Get #1, , CentralDir
          Separate Input(CentralDir.FileLen, 1), Matrix()
          If Len(Matrix(UBound(Matrix))) Then Counter = Counter + 1
          Seek #1, Seek(1) + CentralDir.ExtraLen
          Info(Counter).FileComment = Input(CentralDir.FileCommentLen, 1)
        Case Sign3
          Get #1, , EndCentralDir
          If Len(EndCentralDir.ZipComment) Then
            Load frmComment
            frmComment.txtComment.Text = EndCentralDir.ZipComment
            frmComment.Show vbModal
          End If
          ZipComment = EndCentralDir.ZipComment
          Exit Do
        Case Else
          MsgBox "Unknown signature (0x" + LCase(Hex(Sign)) + ")", vbCritical
          Close
          Exit Sub
      End Select
    Loop
    Close
    .DoNotRefresh = False
    .Refresh
  End With
End Sub

Sub ShowComment()
  If Len(ZipComment) Then
    Load frmComment
    frmComment.txtComment.Text = ZipComment
    frmComment.Show vbModal
  Else
    MsgBox "No comment in zip file", vbInformation
  End If
End Sub

Sub ShowInfo(ItemData As Long)
  Dim Parents() As String, Path As String
  Dim i As Long

  With Info(ItemData)
    If .Nothing Then Exit Sub
    frmMain.txtCompressed.Text = CCur(.CompSize / 1024) & " KB"
    frmMain.txtSize.Text = CCur(.UncompSize / 1024) & " KB"
    frmMain.txtCrc.Text = LCase(Hex(.Crc))
    frmMain.txtRatio.Text = 100 - .CompSize * 100 \ .UncompSize & " %"
    frmMain.usrZipFiles.GetParents ItemData, Parents()
    For i = 1 To UBound(Parents)
      Path = Parents(i) + "\" + Path
    Next i
    frmMain.txtPath.Text = Path
  End With
End Sub

Sub ShowFileInfo(ItemData As Long)
  Dim Parents() As String, Path As String
  Dim i As Long

  With Info(ItemData)
    If .Nothing Then
      MsgBox "Folders do not have properties", vbInformation
      Exit Sub
    End If
    Load frmFileInfo
    frmFileInfo.txtComment.Text = .FileComment
    frmFileInfo.txtExtraField.Text = .ExtraField
    frmFileInfo.txtCompressed.Text = .CompSize & " bytes"
    frmFileInfo.txtSize.Text = .UncompSize & " bytes"
    frmFileInfo.txtCrc.Text = LCase(Hex(.Crc))
    frmFileInfo.txtRatio.Text = 100 - .CompSize * 100 \ .UncompSize & " %"
    frmFileInfo.txtTime.Text = Rest(.Hour) + ":" + Rest(.Minute) + ":" + Rest(.Seconds)
    frmFileInfo.txtDate.Text = .Month & "/" & .Day & "/" & .Year
    frmMain.usrZipFiles.GetParents ItemData, Parents()
    For i = 1 To UBound(Parents)
      Path = Parents(i) + "\" + Path
    Next i
    frmFileInfo.txtPath.Text = Path
    frmFileInfo.Show vbModal
  End With
End Sub

Private Function Rest(ByVal Num As Integer) As String
  Rest = Right("0" & Num, 2)
End Function

Private Sub Separate(Text As String, Matrix() As String)
  Dim i As Long, Max As Long

  ReDim Matrix(0)
  For i = 1 To Len(Text)
    If Mid(Text, i, 1) = "/" Then
      Max = Max + 1
      ReDim Preserve Matrix(Max)
    Else
      Matrix(Max) = Matrix(Max) + Mid(Text, i, 1)
    End If
  Next i
End Sub

Private Sub DateConv(DateVal() As Byte, Mat As UseFullInfo)
  With Mat
    .Month = (DateVal(0) And &HE0) \ &H20
    .Day = DateVal(0) And &H1F
    .Year = DateVal(1) + 1958
  End With
End Sub

Private Sub TimeConv(ByVal TimeVal As Integer, Mat As UseFullInfo)
  With Mat
    .Seconds = (TimeVal And &H1F) * 2
    .Minute = (TimeVal \ &H20) And &H3F
    .Hour = (TimeVal \ &H800) And &H1F
  End With
End Sub
