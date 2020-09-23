VERSION 5.00
Begin VB.UserControl usrTreeView 
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   ScaleHeight     =   339
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   407
   Begin VB.ListBox lstTree 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "usrTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal bytes As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Type ItemInfo
  Id As Long
  Hierarchy As Long
  Parent  As Long
  Text As String
  Opened As Boolean
  Deleted As Boolean
  Childs() As Long
End Type
Dim Separation As Long
Dim Items() As ItemInfo
Dim RaiseInSel As Boolean, NotRefresh As Boolean
Public Event SelectItem(ByVal Changed As Boolean, ByVal ItemId As Long, ByVal ItemText As String, ByVal Hierarchy As Long)

Property Get HierarchySpace() As Long
  HierarchySpace = Separation
End Property

Property Let HierarchySpace(ByVal Sep As Long)
  Separation = Sep
  DrawItems
End Property

Property Get DoNotRefresh() As Boolean
  DoNotRefresh = NotRefresh
End Property

Property Let DoNotRefresh(ByVal Value As Boolean)
  NotRefresh = Value
End Property

Property Get RaiseInSelection() As Boolean
  RaiseInSelection = RaiseInSel
End Property

Property Let RaiseInSelection(ByVal Value As Boolean)
  RaiseInSel = Value
End Property

Private Sub lstTree_DblClick()
  If lstTree.ListIndex = -1 Then Exit Sub
  With Items(lstTree.ItemData(lstTree.ListIndex))
    .Opened = Not .Opened
    If UBound(.Childs) Then DrawItems
    RaiseEvent SelectItem(True, .Id, .Text, .Hierarchy)
  End With
End Sub

Private Sub lstTree_Click() 'MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lstTree.ListIndex = -1 Or Not RaiseInSel Then Exit Sub
  With Items(lstTree.ItemData(lstTree.ListIndex))
    RaiseEvent SelectItem(False, .Id, .Text, .Hierarchy)
  End With
End Sub

Private Sub UserControl_Initialize()
  ReDim Items(0)
  Separation = 4
End Sub

Private Sub UserControl_Resize()
  lstTree.Width = UserControl.ScaleWidth
  lstTree.Height = UserControl.ScaleHeight
End Sub

Private Sub DrawItems()
  Dim Count As Long, Sel As Long
  Dim Backup As Boolean

  If NotRefresh Then Exit Sub
  Backup = RaiseInSel
  RaiseInSel = False
  LockWindowUpdate lstTree.hWnd
  Sel = lstTree.ListIndex
  lstTree.Clear
  For Count = 1 To UBound(Items)
    With Items(Count)
      If .Parent = -1 And Not .Deleted Then
        AddText Items(Count)
        If .Opened Then DrawChilds .Id
      End If
    End With
  Next Count
  If lstTree.ListCount - 1 >= Sel Then
    lstTree.ListIndex = Sel + IIf(lstTree.ListCount - 1 - Sel > 5, 5, lstTree.ListCount - Sel - 1)
    lstTree.ListIndex = Sel
  End If
  LockWindowUpdate 0
  RaiseInSel = Backup
End Sub

Private Sub DrawChilds(ByVal ItemId As Long)
  Dim i As Long

  With Items(ItemId)
    For i = 1 To UBound(.Childs)
      With Items(.Childs(i))
        If Not .Deleted Then
          AddText Items(Items(ItemId).Childs(i))
          If .Opened Then DrawChilds .Id
        End If
      End With
    Next i
  End With
End Sub

Private Sub AddText(Item As ItemInfo)
  Dim TmpStr As String
  Dim i As Long

  With Item
    TmpStr = String(.Hierarchy * Separation, " ") + IIf(UBound(.Childs), IIf(.Opened, "[-]", "[+]"), "[.]") + .Text
    lstTree.AddItem TmpStr
    lstTree.ItemData(lstTree.ListCount - 1) = .Id
  End With
End Sub

Sub Clear()
  lstTree.Clear
  ReDim Items(0)
End Sub

Sub RemoveItem(ByVal ItemId As Long)
  Dim Salir As Boolean
  Dim MaxChild As Long, i As Long

  If ItemId < 0 Or ItemId > UBound(Items) Then
    Err.Raise vbObjectError + 2, , "Item " & ItemId & " not found"
    Exit Sub
  End If
  With Items(ItemId)
    .Deleted = True
    For i = 1 To UBound(Items(.Parent).Childs)
      With Items(.Parent)
        MaxChild = UBound(.Childs)
        If .Childs(i) = ItemId Then
          If MaxChild <> i Then CopyMemory .Childs(i), .Childs(i + 1), (UBound(.Childs) - i - 1) * 4
          Salir = True
        End If
      End With
      If Salir Then Exit For
    Next i
  End With
  ReDim Preserve Items(Items(ItemId).Parent).Childs(MaxChild - 1)
  DrawItems
End Sub

Function AddItem(ByVal ParentId As Long, ByVal Text As String) As Long
  Dim i As Long

  If ParentId < -1 Or ParentId > UBound(Items) Then
    Err.Raise vbObjectError + 1, , "Parent ID " & ParentId & " not found"
    Exit Function
  End If
  ReDim Preserve Items(UBound(Items) + 1)
  With Items(UBound(Items))
    .Id = UBound(Items)
    AddItem = .Id
    .Parent = ParentId
    If ParentId <> -1 Then
      .Hierarchy = Items(ParentId).Hierarchy + 1
      With Items(.Parent)
        ReDim Preserve .Childs(UBound(.Childs) + 1)
        .Childs(UBound(.Childs)) = UBound(Items)
      End With
    End If
    ReDim .Childs(0)
    .Text = Text
  End With
  DrawItems
End Function

Function SearchItem(ByVal Text As String, ByVal Hierarchy As Long, Optional Parent As Long) As Long
  Dim i As Long

  For i = 1 To UBound(Items)
    If Items(i).Text = Text And Items(i).Hierarchy = Hierarchy Then
      SearchItem = Items(i).Id
      Parent = Items(i).Parent
      Exit Function
    End If
  Next i
  SearchItem = -1
  Parent = -1
End Function

Sub Refresh()
  DrawItems
End Sub

Sub SaveTree(ByVal FileName As String)
  Dim FileNum As Integer

  FileNum = FreeFile()
  Open FileName For Binary Access Write As #FileNum
  Put #FileNum, , CLng(UBound(Items))
  Put #FileNum, , Items()
  Close #FileNum
End Sub

Sub LoadTree(ByVal FileName As String)
  Dim FileNum As Integer
  Dim Length As Long

  FileNum = FreeFile()
  Open FileName For Binary Access Read As #FileNum
  Get #FileNum, , Length
  ReDim Items(Length)
  Get #FileNum, , Items()
  Close #FileNum
  DrawItems
End Sub

Sub GetParents(ByVal ItemId As Long, Parents() As String)
  Dim i As Long, Counter As Long

  If ItemId < 1 Or ItemId > UBound(Items) Then
    Err.Raise vbObjectError + 1, , "Item ID " & ItemId & " not found"
    Exit Sub
  End If
  On Error Resume Next
  ReDim Parents(0)
  On Error GoTo 0
  If Err.Number Then
    Err.Raise vbObjectError + 3, , "Dynamic array needed"
    Exit Sub
  End If
  Do
    Counter = Counter + 1
    ItemId = Items(ItemId).Parent
    If ItemId = -1 Then Exit Do
    ReDim Preserve Parents(Counter)
    Parents(Counter) = Items(ItemId).Text
  Loop
End Sub
