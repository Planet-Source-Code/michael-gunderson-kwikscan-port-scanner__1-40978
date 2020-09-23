Attribute VB_Name = "Module1"
'Declare all the functions to use throughout the program
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

Public Const LB_FINDSTRING = &H18F
Public Sub ListKillDupes(listbox As listbox)
'I give full credit to 'source' for the delete duplicates in listbox code
'I used this because it was better than mine. :)
Dim Search1 As Long
Dim Search2 As Long
Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub
Public Function FindInList(lst As listbox, ByVal strItem As String) As Long
'This function uses API function to search a listbox and return the ListIndex #
Dim hList As Long
hList = lst.hwnd
FindInList = SendMessageByString(hList, LB_FINDSTRING, -1, strItem)
End Function
