''
' Global recorder v 1.0
' Michel Verlinden - migul.verlinden@gmail.com - 10/10/2014
'
' Record mouse and keyboard events outside Excel for rapid automation of simple data entry tasks
'
' dependencies: Microsoft Visual Basic for Applications Extensibility 5.3,
'                   Microsoft Forms 2.0 Object Library
' author: migul.verlinden@gmail.com
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
Option Explicit
' API calls
Private Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Private Declare Function GetCursorPos Lib "user32" _
      (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
' Definitions
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Public macroEnd As Boolean
 
' screen coordinates
Type POINTAPI
    X_Pos As Long
    Y_Pos As Long
End Type
 
 
' Main
Sub startRecording()
    Dim cursor As POINTAPI, codeline As String, stopEventTime As Long, _
    startEventTime As Long, waitingTime As Long, currentTextInput As String, _
    textInputEvent As Boolean, i As Integer, keyPress As Integer
    
    ' Initialize time measurements
    stopEventTime = GetTickCount
   
    'Generate new module
    Dim moduleName As String
    moduleName = AddModuleToProject()
   
    Dim switchIdle As Boolean
    switchIdle = False
    textInputEvent = False
    
    Dim stopRecording As Boolean
    stopRecording = False
 
    Dim counter As Integer ' temporary counter for debuggin purposes
    counter = 0
    ' Start listening to mouse and keyboard
    macroEnd = False
    Do While macroEnd = False
        ' test for ENTER key down
        If switchIdle = False And GetAsyncKeyState(vbKeyReturn) Then
            ' Count time since last event
            startEventTime = GetTickCount
            waitingTime = startEventTime - stopEventTime
            Debug.Print "No " & counter & " - Left Click"
            ' If we were typing something the click means we are not in typing mode anymore
            If textInputEvent = True Then
                addOnelineToModule "' " & counter & ": Text input" & _
                                                vbNewLine & "SendKeys """ & currentTextInput & """", moduleName
                currentTextInput = ""
                ' Increment log counter
                counter = counter + 1
                textInputEvent = False
            End If
                addOnelineToModule "' " & counter & ": Text input" & _
                                                vbNewLine & "SendKeys " & """{ENTER}""", moduleName
            ' Flag event
            switchIdle = True
            counter = counter + 1
            Sleep 100
        ' only test if mouse hasn't button hasn't been released
       ' test for LEFT mouse down
        ElseIf switchIdle = False And GetAsyncKeyState(VK_LBUTTON) Then
            ' Count time since last event
            startEventTime = GetTickCount
            waitingTime = startEventTime - stopEventTime
            Debug.Print "No " & counter & " - Left Click"
            ' If we were typing something the click means we are not in typing mode anymore
            If textInputEvent = True Then
                addOnelineToModule "' " & counter & ": Text input" & _
                                                vbNewLine & "SendKeys """ & currentTextInput & """", moduleName
                currentTextInput = ""
                ' Increment log counter
                counter = counter + 1
                textInputEvent = False
            End If
            ' Get Position of cursor on screen
            GetCursorPos cursor
            ' Add line of code to that position
            codeline = set_mouse_and_left_click_code(cursor.X_Pos, cursor.Y_Pos, waitingTime, counter)
            addOnelineToModule codeline, moduleName
            ' Flag event
            switchIdle = True
            counter = counter + 1
            Sleep 100
        ' test for RIGHT mouse button down
        ElseIf switchIdle = False And GetAsyncKeyState(VK_RBUTTON) Then
            ' Count time since last event
            startEventTime = GetTickCount
            waitingTime = stopEventTime - startEventTime
            Debug.Print "No " & counter & " - Right Click"
            ' If we were typing something the click means we are not in typing mode anymore
            If textInputEvent = True Then
                addOnelineToModule "' " & counter & ": Text input" & _
                                                vbNewLine & "SendKeys """ & currentTextInput & """", moduleName
                currentTextInput = ""
                ' Increment log counter
                counter = counter + 1
                textInputEvent = False
            End If
           ' Get Position of cursor on screen
            GetCursorPos cursor
            ' Add line of code to that position
            codeline = set_mouse_and_right_click_code(cursor.X_Pos, cursor.Y_Pos, waitingTime, counter)
            addOnelineToModule codeline, moduleName
            ' Flag event
            switchIdle = True
            counter = counter + 1
            Sleep 100
                ' Test Keyboard input - only in listening mode
        ElseIf switchIdle = False Then
            For i = 1 To 255
                keyPress = 0
                keyPress = GetAsyncKeyState(i)
            'if we find a key that is pressed we attach to our string
                If keyPress = -32767 Then
                    currentTextInput = currentTextInput & Chr$(i)
                    textInputEvent = True
                    ' Event detected - listener should stop until key is released
                    switchIdle = True
                End If
            Next i
        Else
            If switchIdle = True Then
                stopEventTime = GetTickCount
            End If
            ' Re set to listening mode
            switchIdle = False
           
        End If
        ' wait 0.1 seconds for changes and yield to OS
        Sleep 100
        DoEvents
    Loop
    closeSub moduleName
    ThisWorkbook.Save
End Sub
 
'make string to add to module for  left mouse clicks
Function set_mouse_and_left_click_code(x As Long, y As Long, wait As Long, count As Integer) As String
    If count > 0 Then
        set_mouse_and_left_click_code = "' click number " & count & vbNewLine & _
                                "DoEvents" & vbNewLine & _
                                "Sleep " & wait & vbNewLine & _
                                "SetCursorPos " & x & "," & y & vbNewLine & _
                                "mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0" & vbNewLine & _
                                "mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0"
    Else
        set_mouse_and_left_click_code = "'Starting"
    End If
End Function
 
'make string to add to module for right mouse clicks
Function set_mouse_and_right_click_code(x As Long, y As Long, wait As Long, count As Integer) As String
        set_mouse_and_right_click_code = "' click number " & count & vbNewLine & _
                                "DoEvents" & vbNewLine & _
                                "Sleep " & wait & vbNewLine & _
                                "SetCursorPos " & x & "," & y & vbNewLine & _
                                "mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0" & vbNewLine & _
                                "mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0"
End Function
 
' Stop recording
Sub stopRecording()
    macroEnd = True
End Sub
 
 
' Module to add ready to use VBA Windows API code
' Michel Verlinden 10/10/2014
 
' This module has the funtions to write into the VProject
' Code in this module is adapted from Excel MVP C. Pearson examples on
' http://www.cpearson.com/excel/vbe.aspx
'-------------------------------------------------------------
Function AddModuleToProject() As String
    'Access VBE components
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    Set CodeMod = VBComp.CodeModule
    ' Additional module name defined by the current number of modules
    VBComp.Name = "Module" & VBProj.VBComponents.count
    AddModuleToProject = "Module" & VBProj.VBComponents.count
        With CodeMod
            ' Add the API calls necessary to run the recorded macro
            LineNum = .CountOfLines + 1
            .InsertLines LineNum, "Private Declare Sub mouse_event Lib ""user32"" _"
            LineNum = LineNum + 1
            .InsertLines LineNum, "(ByVal dwFlags As Long, ByVal dx As Long, ByVal " & _
                                    "dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Private Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Declare Function SetCursorPos Lib ""user32"" _"
            LineNum = LineNum + 1
            .InsertLines LineNum, "(ByVal x As Long, ByVal y As Long) As Long"
            LineNum = LineNum + 2
            .InsertLines LineNum, "Const MOUSEEVENTF_LEFTDOWN = &H2"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Const MOUSEEVENTF_LEFTUP = &H4"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Const MOUSEEVENTF_MOVE = &H1"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Const MOUSEEVENTF_ABSOLUTE = &H8000"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Const MOUSEEVENTF_RIGHTDOWN = &H8"
            LineNum = LineNum + 1
            .InsertLines LineNum, "Const MOUSEEVENTF_RIGHTUP = &H10"
            LineNum = LineNum + 2
            .InsertLines LineNum, "Sub recorded" & VBProj.VBComponents.count & "()"
        End With
End Function
 
'procedure to add one line of code
Sub addOnelineToModule(line As String, modname As String)
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, CodeMod As VBIDE.CodeModule, _
    LineNum As Long
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(modname)
    Set CodeMod = VBComp.CodeModule
    With CodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, line
    End With
End Sub
 
'procedure to add End Sub after recording
Sub closeSub(modname As String)
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, CodeMod As VBIDE.CodeModule, _
    LineNum As Long
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(modname)
    Set CodeMod = VBComp.CodeModule
    With CodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, "End Sub"
    End With
End Sub
