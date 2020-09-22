Attribute VB_Name = "Module1"
Public Game_Path$


'begin form on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function putMeOnTop(Form As Form)
    SetWindowPos Form.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Function


Public Function takeMeDown(Form As Form)
    SetWindowPos Form.hWnd, -2, 0, 0, 0, 0, 1 Or 2
End Function
'end always on top



