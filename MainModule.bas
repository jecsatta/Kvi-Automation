Attribute VB_Name = "MainModule"
Option Explicit

Public listCommands         As New Dictionary
Public listCommandsItems    As New Dictionary
Public listCommandsDblClick As New Dictionary
Public listStrings         As New Dictionary

Const C_Command_Identifier      As String = "COMMAND:"
Const C_Button_Identifier       As String = "BUTTON:"
Const C_Button_Section          As String = "SECTION:"

Public Const C_Hint_Identifier         As String = "APP_HINT:"
Public Const C_Exit_Identifier         As String = "EXIT_CAPTION:"
Public Const C_Dbl_Click_Identifier    As String = "COMMAND_DBL_CLICK:"
Public Const C_Lang_Identifier         As String = "APP_LANG:"


Public Const C_File_Conf       As String = "\config.txt"
Public Const C_File_Commands   As String = "\commands.txt"
Public Const C_File_Buttons    As String = "\buttons.txt"

 

Public App_ExitCaption As String
Public App_Hint As String
Public App_Lang As String

Private Sub Main()
    If App.PrevInstance Then Exit Sub
    App_Lang = "en-us"
    ReadConfigurations
    ReadLang
    
    App_Hint = "Kvisthor"
    App_ExitCaption = IIf(listStrings.Exists("EXIT"), listStrings.Item("EXIT"), App_ExitCaption)
    
    ReadCommands
    Load frmSysTray
End Sub

Public Sub ExecCommand(vPrefix As String)
    Dim elem As Variant
    For Each elem In listCommandsItems.Keys
        If Left("" & elem, Len(vPrefix) + 1) = vPrefix & "_" Then
            Shell listCommandsItems.Item(elem), vbNormalFocus
        End If
    Next
    Unload frmKVI
End Sub
Public Function FilterDictionaryByKey(originalDict As Dictionary, searchString As String) As Dictionary
    Dim filteredDict As New Dictionary
    Dim currKey As Variant
    For Each currKey In originalDict.Keys
        If InStr(1, UCase(currKey), UCase(searchString)) > 0 Then
            filteredDict.Add currKey, originalDict(currKey)
        End If
    Next currKey

    Set FilterDictionaryByKey = filteredDict
End Function
Private Sub ReadLang()
    Dim vector() As String
    Dim vText As String
    If Dir(App.Path & "\Lang\" & App_Lang & ".txt") = "" Then Exit Sub
    Open App.Path & "\Lang\" & App_Lang & ".txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, vText
        vector = Split(vText, ":", 2)
        listStrings.Add vector(0), vector(1)
    Loop
    Close #1
End Sub

Private Sub ReadCommands()

    Dim vCountCommands   As Long
    Dim vCountCommandItems As Long
    Dim vCountSeparator    As Long
    Dim vLineNumber As Long
    Dim vText      As String
    Dim vCommandName As String
    If Dir(App.Path & C_File_Commands) = "" Then Exit Sub
    Open App.Path & C_File_Commands For Input As #1

    vCountCommands = 0
    vLineNumber = 0
    Do While Not EOF(1)
        vLineNumber = vLineNumber + 1
        Line Input #1, vText
        vText = Trim(vText)
        If Left(vText, 1) <> "#" And vText <> "" Then
            If Left(vText, Len(C_Command_Identifier)) = C_Command_Identifier Then
                vCommandName = Trim(Replace(vText, C_Command_Identifier, ""))
                If vCommandName <> "" Then
                    If vCommandName = "-" Then
                        vCountSeparator = vCountSeparator + 1
                        vCommandName = "Separator" & vCountSeparator
                    End If
                    
                    If Not listCommands.Exists(vCommandName) Then
                        vCountCommands = vCountCommands + 1
                        listCommands.Add vCommandName, vCountCommands
                        vCountCommandItems = 0
                    Else
                        MsgBox listStrings.Item("COMMAND_ALREADY_EXISTS") & vLineNumber
                        End
                    End If
                Else
                    MsgBox listStrings.Item("COMMAND_WITHOUT_NAME") & vLineNumber
                    End
                End If
            Else
                If listCommands.Count > 0 Then
                    vCountCommandItems = vCountCommandItems + 1
                    listCommandsItems.Add vCountCommands & "_" & vCountCommandItems, Trim(vText)
                Else
                    MsgBox listStrings.Item("COMMAND_WITHOUT_PARENT") & vLineNumber
                    End
                End If
            End If
        End If
    Loop
 Close #1
End Sub

Private Sub ReadConfigurations()
    Dim vLineNumber As Long
    Dim vCountDblClickCommands As Long
    Dim vText      As String
    If Dir(App.Path & C_File_Conf) = "" Then Exit Sub
    Open App.Path & C_File_Conf For Input As #1
    
    vLineNumber = 0
    vCountDblClickCommands = 0
    Do While Not EOF(1)
        vLineNumber = vLineNumber + 1
        Line Input #1, vText
        vText = Trim(vText)
        If Left(vText, 1) <> "#" And vText <> "" Then
            If Left(vText, Len(C_Hint_Identifier)) = C_Hint_Identifier Then
                App_Hint = Trim(Replace(vText, C_Hint_Identifier, ""))
            ElseIf Left(vText, Len(C_Exit_Identifier)) = C_Exit_Identifier Then
                App_ExitCaption = Trim(Replace(vText, C_Exit_Identifier, ""))
            ElseIf Left(vText, Len(C_Dbl_Click_Identifier)) = C_Dbl_Click_Identifier Then
                vCountDblClickCommands = vCountDblClickCommands + 1
                listCommandsDblClick.Add vCountDblClickCommands, Trim(Replace(vText, C_Dbl_Click_Identifier, ""))
            ElseIf Left(vText, Len(C_Lang_Identifier)) = C_Lang_Identifier Then
                App_Lang = Trim(Replace(vText, C_Lang_Identifier, ""))
            End If
        End If
    Loop
 Close #1
End Sub
