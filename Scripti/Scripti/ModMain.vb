﻿Module ModMain
    Public debugMode As Boolean = False
    Public exitScript As Boolean = False
    Public scripts As New List(Of IO.FileInfo)
    Public strings As New Dictionary(Of String, String)
    Public integers As New Dictionary(Of String, Integer)
    Public booleans As New Dictionary(Of String, Boolean)
    Private ignoreVersion As Boolean = False
    Private wait As Boolean = False

    Public Enum YesNoQuestionDefault
        [Nothing]
        Yes
        No
    End Enum

    Sub Main()
        With Initialize()
            If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                ShowError(.Error)
            End If
        End With

        For Each script In scripts
            HLine()
            Console.WriteLine(" Script: " & script.ToString)
            HLine()
            With InterpretScript(script)
                If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                    ShowError(.Error)
                End If
            End With
            HLine()
            Console.WriteLine()
        Next

        If wait = True Then
            Pause()
        End If
    End Sub

    ''' <summary>
    ''' Initalizes Scripti
    ''' </summary>
    ''' <returns>Success or an error</returns>
    Function Initialize() As ScriptEventInfo
        Try
            Console.Title = "Scripti Version " & My.Application.Info.Version.ToString
            Console.WriteLine("Scripti Version " & My.Application.Info.Version.ToString)
            Console.WriteLine("Copyright (c) 2015 Mario Wagenknecht")
            Console.WriteLine("Source at: https://github.com/master-m1000/Scripti")
            Console.WriteLine()
            For Each argument In Environment.GetCommandLineArgs
                Select Case argument
                    Case Environment.GetCommandLineArgs(0)
                    Case "debug"
                        debugMode = (argument = "debug")
                        Console.WriteLine("DEBUG MODE")
                        Console.WriteLine()
                    Case "wait"
                        wait = True
                    Case Else
                        With AddScript(argument)
                            If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                                ShowError(.Error)
                            End If
                        End With
                End Select
            Next
            Console.WriteLine("List of script files:")
            For Each script In scripts
                Console.WriteLine(" " & script.ToString)
            Next
            Console.WriteLine()

            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function



    ''' <summary>
    ''' Interprets a script
    ''' </summary>
    ''' <param name="script"></param>
    ''' <returns>Success or an error</returns>
    Function InterpretScript(ByVal script As IO.FileInfo) As ScriptEventInfo
        Dim lineNo As Integer = 0
        Try
            exitScript = False
            For Each line In IO.File.ReadAllLines(script.ToString)
                lineNo += 1
                If lineNo = 1 Then
                    If line.StartsWith("|SCRIPTI SCRIPT FILE VERSION ") = True Then
                        If Not line.Replace("|SCRIPTI SCRIPT FILE VERSION ", "") = My.Application.Info.Version.ToString And ignoreVersion = False Then
                            Console.WriteLine("This script is not written for this version of Scripti.")
                            If AskYesNo("Do you want to continue?", YesNoQuestionDefault.No) = False Then
                                Exit For
                            End If
                        End If
                    Else
                        Throw New Exception("Not a valid script file.")
                    End If
                Else
                    With InterpretLine(line)
                        If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                            ShowError(.Error, lineNo)
                        End If
                    End With
                End If
                If exitScript = True Then
                    Exit For
                End If
            Next
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Shows an error
    ''' </summary>
    ''' <param name="ex">Exception</param>
    ''' <returns>Success or an error</returns>
    Function ShowError(ByVal ex As Exception, Optional ByVal lineNo As Integer = -1) As ScriptEventInfo
        Try
            If lineNo = -1 Then
                If debugMode = True Then
                    Console.WriteLine("ERROR: " & ex.ToString)
                Else
                    Console.WriteLine("ERROR: " & ex.Message)
                End If
            Else
                If debugMode = True Then
                    Console.WriteLine("ERROR Line " & lineNo & ": " & ex.ToString)
                Else
                    Console.WriteLine("ERROR Line " & lineNo & ": " & ex.Message)
                End If
            End If
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex2 As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex2}
        End Try
    End Function

    ''' <summary>
    ''' Wait for a pressed key until continue
    ''' </summary>
    ''' <returns>Success or an error</returns>
    Function Pause() As ScriptEventInfo
        Try
            Console.Write("Press any key to continue . . . ")
            Console.ReadKey(True)
            Console.WriteLine()
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Asks a Yes/No-Question
    ''' </summary>
    ''' <param name="question">A Yes/No-Question</param>
    ''' <param name="defaultAnswer">Default Answer if enter pressed</param>
    ''' <returns></returns>
    Function AskYesNo(ByVal question As String, Optional ByVal defaultAnswer As YesNoQuestionDefault = YesNoQuestionDefault.Nothing) As Boolean
        Dim answer As Boolean
        While True
            Select Case defaultAnswer
                Case YesNoQuestionDefault.Nothing
                    Console.Write(question & " [y/n] ")
                Case YesNoQuestionDefault.Yes
                    Console.Write(question & " [Y/n] ")
                Case YesNoQuestionDefault.No
                    Console.Write(question & " [y/N] ")
            End Select
            Dim x As ConsoleKeyInfo = Console.ReadKey
            Select Case x.Key
                Case ConsoleKey.Y
                    answer = True
                    Exit While
                Case ConsoleKey.N
                    answer = False
                    Exit While
                Case ConsoleKey.Enter
                    Select Case defaultAnswer
                        Case YesNoQuestionDefault.Yes
                            answer = True
                            Exit While
                        Case YesNoQuestionDefault.No
                            answer = False
                            Exit While
                        Case YesNoQuestionDefault.Nothing
                            answer = Nothing
                    End Select
                Case Else
                    answer = Nothing
            End Select
            Console.WriteLine()
        End While
        Console.WriteLine()
        Return answer
    End Function

    ''' <summary>
    ''' Draws a horizontal line.
    ''' </summary>
    ''' <returns></returns>
    Function HLine() As ScriptEventInfo
        Try
            For i As Integer = 1 To Console.WindowWidth
                Console.Write("-")
            Next
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Adds a script to the script files list
    ''' </summary>
    ''' <param name="path">Path of the script file</param>
    ''' <returns>Success or an error</returns>
    Function AddScript(ByVal path As String) As ScriptEventInfo
        Try
            Dim script As New IO.FileInfo(path)
            If script.Exists = True Then
                scripts.Add(script)
            Else
                Throw New Exception("Could not find script: " & script.ToString)
            End If
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    'Templates
    '
    '''' <summary>
    '''' This is a template
    '''' </summary>
    '''' <returns>Success or an error</returns>
    'Function Template() As ScriptEventInfo
    '    Try
    '        'Do something
    '        Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
    '    Catch ex As Exception
    '        Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
    '    End Try
    'End Function
    '
    'With Template()
    'If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
    '            ShowError(.Error)
    '        End If
    'End With

End Module