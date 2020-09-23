VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
' ListView Sorting
'-------------------------------------------------------------------------------

Private m_blnDirection(1 To 3) As Boolean   ' Sort Direction for each column
                                            ' in the ListView

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView ListView1, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView ListView1, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView ListView1, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    End Select
End Sub

'*******************************************************************************
' Test Harness Infrastructure
'-------------------------------------------------------------------------------

Private Sub Form_Load()
    With ListView1
    
        ' Set ListView Properties
        
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "String"
        .ColumnHeaders.Add , , "Number"
        .ColumnHeaders.Add , , "Date"
        
        ' Populate the ListView with Junk
        
        Dim i As Integer
        For i = 1 To 100
            With .ListItems.Add(, , RandomString)
                .ListSubItems.Add , , RandomNumber
                .ListSubItems.Add , , RandomDate
            End With
        Next
    End With
End Sub

Private Sub Form_Resize()
    ' Size the ListView with the form
    ListView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Function RandomString() As String
    Const CHARS As Integer = 3
    Dim i As Integer, str As String
    For i = 1 To CHARS
        str = str & Chr$(Asc("A") + CInt(Rnd * 25))
    Next
    RandomString = str
End Function

Private Function RandomNumber() As String
    Const RANGE As Integer = 200
    RandomNumber = Format$((Rnd * RANGE) - (RANGE / 2), "0.00")
End Function

Private Function RandomDate() As String
    Const RANGE As Integer = 200
    RandomDate = Format$(DateAdd("d", CInt(Rnd * RANGE) - (RANGE / 2), Date), _
                                                                "DD/MM/YYYY")
End Function

'*******************************************************************************
'
'-------------------------------------------------------------------------------

