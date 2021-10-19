'Displaing messages on Excel statusbar

Sub statusBarExample()
    Application.DisplayStatusBar = True 

    Application.StatusBar = "text"

    Application.StatusBar = False 

End Sub

Sub changeCursorExample()
    Application.Cursor = xlWait

    Application.Cursor = xlDefault 
End Sub
