<div align="center">

## Sort a ListView by Number or Date \(Updated\)


</div>

### Description

This code allows a ListView control to be sorted by Number or Date without having to use APIs (except to lock the screen)
 
### More Info
 
No known side-effects at this time. Does not mess-up the

ListItems collection like a Custom API-implemented sort.


<span>             |<span>
---                |---
**Submitted On**   |1999-12-14 10:03:36
**By**             |[Pete Cozens](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pete-cozens.md)
**Level**          |Advanced
**User Rating**    |4.8 (58 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD229112141999\.zip](https://github.com/Planet-Source-Code/pete-cozens-sort-a-listview-by-number-or-date-updated__1-1855/archive/master.zip)





### Source Code

```
'****************************************************************
' ListView1_ColumnClick
' Called when a column header is clicked on - sorts the data in
' that column
'----------------------------------------------------------------
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As _
                  MSComctlLib.ColumnHeader)
  On Error Resume Next
  ' Record the starting CPU time (milliseconds since boot-up)
  Dim lngStart As Long
  lngStart = GetTickCount
  ' Commence sorting
  With ListView1
    ' Display the hourglass cursor whilst sorting
    Dim lngCursor As Long
    lngCursor = .MousePointer
    .MousePointer = vbHourglass
    ' Prevent the ListView control from updating on screen -
    ' this is to hide the changes being made to the listitems
    ' and also to speed up the sort
    LockWindowUpdate .hWnd
    ' Check the data type of the column being sorted,
    ' and act accordingly
    Dim l As Long
    Dim strFormat As String
    Dim strData() As String
    Dim lngIndex As Long
    lngIndex = ColumnHeader.Index - 1
    Select Case UCase$(ColumnHeader.Tag)
    Case "DATE"
      ' Sort by date.
      strFormat = "YYYYMMDDHhNnSs"
      ' Loop through the values in this column. Re-format
      ' the dates so as they can be sorted alphabetically,
      ' having already stored their visible values in the
      ' tag, along with the tag's original value
      With .ListItems
        If (lngIndex > 0) Then
          For l = 1 To .Count
            With .Item(l).ListSubItems(lngIndex)
              .Tag = .Text & Chr$(0) & .Tag
              If IsDate(.Text) Then
                .Text = Format(CDate(.Text), _
                          strFormat)
              Else
                .Text = ""
              End If
            End With
          Next l
        Else
          For l = 1 To .Count
            With .Item(l)
              .Tag = .Text & Chr$(0) & .Tag
              If IsDate(.Text) Then
                .Text = Format(CDate(.Text), _
                          strFormat)
              Else
                .Text = ""
              End If
            End With
          Next l
        End If
      End With
      ' Sort the list alphabetically by this column
      .SortOrder = (.SortOrder + 1) Mod 2
      .SortKey = ColumnHeader.Index - 1
      .Sorted = True
      ' Restore the previous values to the 'cells' in this
      ' column of the list from the tags, and also restore
      ' the tags to their original values
      With .ListItems
        If (lngIndex > 0) Then
          For l = 1 To .Count
            With .Item(l).ListSubItems(lngIndex)
              strData = Split(.Tag, Chr$(0))
              .Text = strData(0)
              .Tag = strData(1)
            End With
          Next l
        Else
          For l = 1 To .Count
            With .Item(l)
              strData = Split(.Tag, Chr$(0))
              .Text = strData(0)
              .Tag = strData(1)
            End With
          Next l
        End If
      End With
    Case "NUMBER"
      ' Sort Numerically
      strFormat = String(30, "0") & "." & String(30, "0")
      ' Loop through the values in this column. Re-format the values so as they
      ' can be sorted alphabetically, having already stored their visible
      ' values in the tag, along with the tag's original value
      With .ListItems
        If (lngIndex > 0) Then
          For l = 1 To .Count
            With .Item(l).ListSubItems(lngIndex)
              .Tag = .Text & Chr$(0) & .Tag
              If IsNumeric(.Text) Then
                If CDbl(.Text) >= 0 Then
                  .Text = Format(CDbl(.Text), _
                    strFormat)
                Else
                  .Text = "&" & InvNumber( _
                    Format(0 - CDbl(.Text), _
                    strFormat))
                End If
              Else
                .Text = ""
              End If
            End With
          Next l
        Else
          For l = 1 To .Count
            With .Item(l)
              .Tag = .Text & Chr$(0) & .Tag
              If IsNumeric(.Text) Then
                If CDbl(.Text) >= 0 Then
                  .Text = Format(CDbl(.Text), _
                    strFormat)
                Else
                  .Text = "&" & InvNumber( _
                    Format(0 - CDbl(.Text), _
                    strFormat))
                End If
              Else
                .Text = ""
              End If
            End With
          Next l
        End If
      End With
      ' Sort the list alphabetically by this column
      .SortOrder = (.SortOrder + 1) Mod 2
      .SortKey = ColumnHeader.Index - 1
      .Sorted = True
      ' Restore the previous values to the 'cells' in this
      ' column of the list from the tags, and also restore
      ' the tags to their original values
      With .ListItems
        If (lngIndex > 0) Then
          For l = 1 To .Count
            With .Item(l).ListSubItems(lngIndex)
              strData = Split(.Tag, Chr$(0))
              .Text = strData(0)
              .Tag = strData(1)
            End With
          Next l
        Else
          For l = 1 To .Count
            With .Item(l)
              strData = Split(.Tag, Chr$(0))
              .Text = strData(0)
              .Tag = strData(1)
            End With
          Next l
        End If
      End With
    Case Else  ' Assume sort by string
      ' Sort alphabetically. This is the only sort provided
      ' by the MS ListView control (at this time), and as
      ' such we don't really need to do much here
      .SortOrder = (.SortOrder + 1) Mod 2
      .SortKey = ColumnHeader.Index - 1
      .Sorted = True
    End Select
    ' Unlock the list window so that the OCX can update it
    LockWindowUpdate 0&
    ' Restore the previous cursor
    .MousePointer = lngCursor
  End With
  ' Report time elapsed, in milliseconds
  Debug.Print "Time Elapsed = " & GetTickCount - lngStart & "ms"
End Sub
'****************************************************************
' InvNumber
' Function used to enable negative numbers to be sorted
' alphabetically by changing the characters
'----------------------------------------------------------------
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
'****************************************************************
'
'----------------------------------------------------------------
```

