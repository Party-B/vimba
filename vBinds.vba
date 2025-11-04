Option Explicit
Public vimMode As String
Public sBufferBar As String

Function BindKeys(ByVal keyCode As Long, ByVal macroName As String)
    On Error Resume Next
    
    CustomizationContext = NormalTemplate
    
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:=macroName, _
                    keyCode:=keyCode
    
    If Err.Number <> 0 Then
        MsgBox "What error?: " & Err.Description
        Err.Clear
    Else
        ' Debug.Print keyCode & " bound to " & macroName
    End If
    
End Function

Sub SetMode(whichMode As String)
    
    Select Case whichMode
    
        Case "NORMAL"
            vimMode = whichMode
            Application.StatusBar = Space(50) & "-- " & whichMode & " --"
            sBufferBar = ""
            
        Case "INSERT"
            vimMode = whichMode
            Application.StatusBar = Space(50) & "-- " & whichMode & " --"
            
        Case "BUFFER"
            vimMode = whichMode
            
        Case "VISUAL"
            vimMode = whichMode
            Application.StatusBar = Space(50) & "-- " & whichMode & " --"
            
    End Select

End Sub


Function BarBuffer(whichKey As String)
    
    sBufferBar = sBufferBar & whichKey
    
        Application.StatusBar = Space(50) & sBufferBar
    
    If isActionKey(whichKey) Then
        'Carry out action with mod
        Dim ThisAction As String
        Dim xTimes As Integer
        Dim numStr As String
         
        ThisAction = whichKey
       ' TODO: Need to fix the buffer now because we originally relied on having : all the time.
       
        numStr = Left(sBufferBar, InStr(1, sBufferBar, ThisAction) - 1) ' everything before the key
        
        If isNumber(numStr) Then xTimes = CInt(numStr)
        
        SetMode "NORMAL"
        
        DoAction ThisAction, xTimes
        
        sBufferBar = ""
    End If
    
End Function

' === INSERTS ====
' ################

Public Sub localKeyI()
    If vimMode = "INSERT" Then
        Selection.TypeText "i"
        Exit Sub
    ElseIf vimMode = "NORMAL" Then
        SetMode "INSERT"
    ElseIf vimMode = "BUFFER" Then
        
    End If
    
End Sub

Public Sub localKeyA()
    If vimMode = "INSERT" Then
        Selection.TypeText "a"
        Exit Sub
    ElseIf vimMode = "NORMAL" Then
        SetMode "INSERT"
    ElseIf vimMode = "BUFFER" Then
        
    End If
    
End Sub


' === MOVEMENT ===
' ################

Public Sub localKeyH()
    If vimMode = "INSERT" Then
        Selection.TypeText "h"
    ElseIf vimMode = "NORMAL" Then
        Selection.MoveLeft wdCharacter, 1
    ElseIf vimMode = "BUFFER" Then
        
    End If

End Sub
Public Sub localKeyJ()
    If vimMode = "INSERT" Then
        Selection.TypeText "j"
    ElseIf vimMode = "NORMAL" Then
        Selection.MoveDown Unit:=wdLine, Count:=1
    ElseIf vimMode = "BUFFER" Then
        
    End If

End Sub
Public Sub localKeyK()
    If vimMode = "INSERT" Then
        Selection.TypeText "K"
    ElseIf vimMode = "NORMAL" Then
        Selection.MoveUp Unit:=wdLine, Count:=1
    ElseIf vimMode = "BUFFER" Then
        
    End If

End Sub
Public Sub localKeyL()
    If vimMode = "INSERT" Then
        Selection.TypeText "l"
    ElseIf vimMode = "NORMAL" Then
        Selection.MoveRight wdCharacter, 1
    ElseIf vimMode = "BUFFER" Then
        
    End If

End Sub

' === WORDS ===
' Got to remember to diff W from w.

Public Sub localKeyW(Optional sMod As String)
    
    If sMod = "" Then sMod = 1

    If vimMode = "INSERT" Then
        Selection.TypeText "w"
    ElseIf vimMode = "NORMAL" Then
        Selection.MoveRight wdWord, CInt(sMod)
    ElseIf vimMode = "BUFFER" Then
        BarBuffer "w"
    End If

End Sub

Public Sub localKeyE(Optional sMod As String)
    If sMod = "" Then sMod = 1

    If vimMode = "INSERT" Then
        Selection.TypeText "e"
    ElseIf vimMode = "NORMAL" Then
        ' Check if we on a space
        Selection.MoveRight wdCharacter, 1, wdExtend
        If Selection.Text = " " Then
            Selection.Collapse wdCollapseEnd
        Else
            Selection.Collapse wdCollapseStart
        End If
    
        Selection.MoveRight Unit:=wdWord, Count:=CInt(sMod), Extend:=wdMove
        
        ' Check if we on a space again
        Selection.MoveLeft wdCharacter, 1, wdExtend
        If Selection.Text = " " Then
            Selection.Collapse wdCollapseStart
        Else
            Selection.Collapse wdCollapseEnd
        End If
        
    ElseIf vimMode = "BUFFER" Then
        BarBuffer "e"
    End If

End Sub

Public Sub localKeyB(Optional sMod As String)
    If sMod = "" Then sMod = 1
    
    If vimMode = "INSERT" Then
        Selection.TypeText "b"
    ElseIf vimMode = "NORMAL" Then
        ' Check if we on a space
        Selection.MoveLeft wdCharacter, 1, wdExtend
        If Selection.Text = " " Then
            Selection.Collapse wdCollapseStart
        Else
            Selection.Collapse wdCollapseEnd
        End If
    
        Selection.MoveLeft Unit:=wdWord, Count:=CInt(sMod), Extend:=wdMove
        
        ' Check if we on a space again
        Selection.MoveRight wdCharacter, 1, wdExtend
        If Selection.Text = " " Then
            Selection.Collapse wdCollapseEnd
        Else
            Selection.Collapse wdCollapseStart
        End If
        
    ElseIf vimMode = "BUFFER" Then
        
    End If

End Sub

' === DELETE ===
' ##############

Public Sub localKeyD(Optional sMod As String)
    ' UP TO HERE:
    ' This is a bit of a pain, given that we implemented the buffer
    ' but this is kind of a second buffer that doesn't need the entry
    ' with the colon.
    If sMod = "" Then sMod = 1

    If vimMode = "INSERT" Then
        Selection.TypeText "d"
    ElseIf vimMode = "NORMAL" Then
        SetMode "BUFFER"
        BarBuffer "d"
    ElseIf vimMode = "BUFFER" Then
        BarBuffer "d"
    End If

End Sub

' === BUFFER ===
' ##############

Public Sub localKeyColon()
    If vimMode = "INSERT" Then
        Selection.TypeText ":"
    ElseIf vimMode = "NORMAL" Then
        SetMode "BUFFER"
        BarBuffer ":"
    End If

End Sub


' === NUMBERS ===
' ###############

Public Sub localKey1()
    If vimMode = "INSERT" Then
        Selection.TypeText "1"
    ElseIf vimMode = "NORMAL" Then
        ' TODO:
        
    ElseIf vimMode = "BUFFER" Then
        BarBuffer "1"
    End If
End Sub

Public Sub localKey2()
    If vimMode = "INSERT" Then
        Selection.TypeText "2"
    ElseIf vimMode = "NORMAL" Then
        ' TODO:
        
    ElseIf vimMode = "BUFFER" Then
        BarBuffer "2"
    End If
End Sub

Function isActionKey(whichKey As String) As Boolean
       
    Select Case whichKey
    
        Case "w"
            isActionKey = True
        
        Case "e"
            isActionKey = True
            
        Case "b"
            isActionKey = True
        
        Case Else
            isActionKey = False
            
    End Select
            
End Function

Sub DoAction(whichKey As String, xTimes As Integer)

    Select Case whichKey
    
        Case "w"
            ' Move word
            localKeyW CStr(xTimes)
        Case "e"
            ' Move word
            localKeyE CStr(xTimes)
        Case "b"
            ' Move word
            localKeyB CStr(xTimes)
    End Select
End Sub

' ===== GENERAL FUNCTIONS =====
    
    Function isNumber(whichString As String) As Boolean

    Dim i As Integer
    
    For i = 1 To Len(whichString)
    
        If IsNumeric(Mid(whichString, i, 1)) Then
            isNumber = True
        Else
            isNumber = False
            Exit Function
        End If
    Next
    
End Function

