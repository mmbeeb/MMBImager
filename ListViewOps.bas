Attribute VB_Name = "ListViewOps"
' General ListView operations
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

Public Sub LVSelectAll(lv As ListView)
    ' ListView: Select everything
    Dim x As Integer
    For x = 1 To lv.ListItems.Count
        lv.ListItems(x).Selected = True
    Next
End Sub

Public Sub LVSelectNothing(lv As ListView)
    ' ListView: Select nothing
    Dim x As Integer
    Set lv.SelectedItem = Nothing
    For x = 1 To lv.ListItems.Count
        lv.ListItems(x).Selected = False
    Next
End Sub

Public Function LVSelectCount(lv As ListView) As Integer
    ' ListView: Return no. of selected items
    Dim x As Integer
    Dim c As Integer
    For x = 1 To lv.ListItems.Count
        If lv.ListItems(x).Selected Then c = c + 1
    Next
    LVSelectCount = c
End Function

Public Function LVIndexUnderPoint(lv As ListView, x As Single, y As Single) As Long
    ' Return index of item under point
    If lv.HitTest(x, y) Is Nothing Then
        LVIndexUnderPoint = 0
    Else
        LVIndexUnderPoint = lv.HitTest(x, y).Index
    End If
End Function
