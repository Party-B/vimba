Option Explicit

Public bVimbaOn As Boolean

Sub ToggleVimba()
    
    If bVimbaOn Then
        ClearBindings
    Else
        RunBindings
    End If
   
End Sub

Sub RunBindings()
    
    ' Initialise the environment
    bVimbaOn = True
    vBinds.SetMode "NORMAL"
    
    ' Run the binds = yeah, I know it's stupid but keybinds in word is touchy.
    
    
    vBinds.BindKeys wdKeyI, "VimbaKeyI"
    vBinds.BindKeys wdKeyEsc, "EscapeMode"
    
    ' === BUFFER ===
    ' ##############
    vBinds.BindKeys wdKeyShift + wdKeySemiColon, "VimbaKeycolon"
    vBinds.BindKeys wdKeyD, "VimbaKeyd"
    
    
    ' === MOVEMENT ===
    ' ################

    vBinds.BindKeys wdKeyH, "VimbaKeyh"
    vBinds.BindKeys wdKeyJ, "VimbaKeyj"
    vBinds.BindKeys wdKeyK, "VimbaKeyk"
    vBinds.BindKeys wdKeyL, "VimbaKeyl"
    
    ' === NAVIGATION ===
    ' ##################
    
    vBinds.BindKeys wdKeyW, "VimbaKeyw"
    vBinds.BindKeys wdKeyE, "VimbaKeye"
    vBinds.BindKeys wdKeyB, "VimbaKeyb"
    
    ' === NUMBERS ===
    ' ###############
    
    vBinds.BindKeys wdKey1, "VimbaKey1"
    vBinds.BindKeys wdKey2, "VimbaKey2"
    
    MsgBox "Vimba mode on!", vbCritical, "Vimba"

End Sub

Sub ClearBindings()
    On Error Resume Next
    
    Dim kb As keyBinding
    For Each kb In KeyBindings
        
        kb.Clear
 
    Next
    bVimbaOn = False
    MsgBox "Bindings cleared"
End Sub

Sub EnterInsertMode()
    If vimMode = "NORMAL" Then
        SetMode "INSERT"
    End If
End Sub

Sub EscapeMode()
    SetMode "NORMAL"
End Sub

' ==== BINDS MUST BE SET THROUGH HERE ====
' TO POINT TO LOCAL WHERE MORE IS DONE.
Sub VimbaKey1()
    vBinds.localKey1
End Sub

Sub VimbaKey2()
    vBinds.localKey2
End Sub
Sub VimbaKeya()
    vBinds.localKeyA
End Sub

Sub VimbaKeyb()
    vBinds.localKeyB
End Sub
Sub VimbaKeyd()
    vBinds.localKeyD
End Sub
Sub VimbaKeye()
    vBinds.localKeyE
End Sub

Sub VimbaKeyh()
    vBinds.localKeyH
End Sub

Sub VimbaKeyi()
    vBinds.localKeyI
End Sub

Sub VimbaKeyj()
    vBinds.localKeyJ
End Sub

Sub VimbaKeyk()
    vBinds.localKeyK
End Sub

Sub VimbaKeyl()
    vBinds.localKeyL
End Sub

Sub VimbaKeyw()
    vBinds.localKeyW
End Sub

Sub VimbaKeycolon()
    vBinds.localKeyColon
End Sub



