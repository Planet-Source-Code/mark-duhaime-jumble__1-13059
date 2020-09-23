Attribute VB_Name = "Module1"

Sub OpenSkin(filename As String)
    Dim no As Integer
    Dim i As Integer
    
    On Error Resume Next ' In case of error

    no = FreeFile
    Open filename For Input As #no
    Do
        Input #no, A
        Input #no, b
        Input #no, c
        Input #no, d
        Input #no, e
        Input #no, f
        Input #no, g
        Input #no, h
        Input #no, i
        Input #no, J
        Input #no, K
        Input #no, l
        Input #no, m
        If b = "[Skin]" Then
            If c = "" Then
                '
            Else ' else continue loading skin
            'Load images (c is the path name)
                frmJumble.Picture = LoadPicture(App.Path + "\" + d) ' This will load the picture from the path you choose in the skin editor
                frmJumble.Mask.Picture = LoadPicture(App.Path + "\" + e)  ' This will load the mask from the path you choose in the skin editor
                ' Open size settings
                frmJumble.Height = f
                frmJumble.MoveForm.Height = frmJumble.Height ' the move picture label
                frmJumble.Width = g
                frmJumble.MoveForm.Width = frmJumble.Width ' the move picture label
            End If
        End If
    Loop Until EOF(no)
    Close #no ' close file
    Call frmJumble.ChangeMask ' This is for updating the mask
End Sub
