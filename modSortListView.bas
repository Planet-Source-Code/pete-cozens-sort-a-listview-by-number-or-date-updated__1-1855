Attribute VB_Name = "modSortListView"
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Enum ListDataType
    ldtString = 0
    ldtNumber = 1
    ldtDateTime = 2
End Enum

'*******************************************************************************
' Sort a ListView by String, Number, or DateTime
'
' Parameters:
'
'   ListView    Reference to the ListView control to be sorted.
'   Index       Index of the column in the ListView to be sorted. The first
'               column in a ListView has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

Public Sub SortListView(ListView As ListView, ByVal Index As Integer, _
                ByVal DataType As ListDataType, ByVal Ascending As Boolean)

    On Error Resume Next
    Dim i As Integer
    Dim l As Long
    Dim strFormat As String
    
    ' Display the hourglass cursor whilst sorting
    
    Dim lngCursor As Long
    lngCursor = ListView.MousePointer
    ListView.MousePointer = vbHourglass
    
    ' Prevent the ListView control from updating on screen - this is to hide
    ' the changes being made to the listitems, and also to speed up the sort
    
    LockWindowUpdate ListView.hWnd
    
    Dim blnRestoreFromTag As Boolean
    
    Select Case DataType
    Case ldtString
        
        ' Sort alphabetically. This is the only sort provided by the
        ' MS ListView control (at this time), and as such we don't really
        ' need to do much here
    
        blnRestoreFromTag = False
        
    Case ldtNumber
    
        ' Sort Numerically
    
        strFormat = String$(20, "0") & "." & String$(10, "0")
        
        ' Loop through the values in this column. Re-format the values so
        ' as they can be sorted alphabetically, having already stored their
        ' text values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        If IsNumeric(.Text) Then
                            If CDbl(.Text) >= 0 Then
                                .Text = Format(CDbl(.Text), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                            End If
                        Else
                            .Text = ""
                        End If
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        If IsNumeric(.Text) Then
                            If CDbl(.Text) >= 0 Then
                                .Text = Format(CDbl(.Text), strFormat)
                            Else
                                .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                            End If
                        Else
                            .Text = ""
                        End If
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
    
    Case ldtDateTime
    
        ' Sort by date.
        
        strFormat = "YYYYMMDDHhNnSs"
        
        Dim dte As Date
    
        ' Loop through the values in this column. Re-format the dates so as they
        ' can be sorted alphabetically, having already stored their visible
        ' values in the tag, along with the tag's original value
    
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(Index - 1)
                        .Tag = .Text & Chr$(0) & .Tag
                        dte = CDate(.Text)
                        .Text = Format$(dte, strFormat)
                    End With
                Next l
            End If
        End With
        
        blnRestoreFromTag = True
        
    End Select
    
    ' Sort the ListView Alphabetically
    
    ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
    ListView.SortKey = Index - 1
    ListView.Sorted = True
    
    ' Restore the Text Values if required
    
    If blnRestoreFromTag Then
        
        ' Restore the previous values to the 'cells' in this column of the list
        ' from the tags, and also restore the tags to their original values
        
        With ListView.ListItems
            If (Index = 1) Then
                For l = 1 To .Count
                    With .Item(l)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            Else
                For l = 1 To .Count
                    With .Item(l).ListSubItems(Index - 1)
                        i = InStr(.Tag, Chr$(0))
                        .Text = Left$(.Tag, i - 1)
                        .Tag = Mid$(.Tag, i + 1)
                    End With
                Next l
            End If
        End With
    End If
    
    ' Unlock the list window so that the OCX can update it
    
    LockWindowUpdate 0&
    
    ' Restore the previous cursor
    
    ListView.MousePointer = lngCursor
    

End Sub

'*******************************************************************************
' Modifies a numeric string to allow it to be sorted alphabetically
'-------------------------------------------------------------------------------

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

'*******************************************************************************
'
'-------------------------------------------------------------------------------

