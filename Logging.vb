' Name:     Logging
' Author:   Donald G Plugge
' Date:     7/16/07
' Purpose:  Class for logging activity
Imports System.io
Imports FCS_Classes.Utility
Public Class Logging

    Public Logging_Flag As Boolean
    Public TS_log As System.IO.StreamWriter
    Public App_Log_File As String

    Private fso As New Scripting.FileSystemObject

    Private mPrefix As String
    Public Property Prefix() As String
        Get
            Return mPrefix
        End Get
        Set(ByVal value As String)
            mPrefix = value
        End Set
    End Property
    ' dgp rev 7/16/07 Reopen the log file
    Public Sub Reopen()

        TS_log = New StreamWriter(App_Log_File, True)

    End Sub
    ' dgp rev 7/16/07 Reset the log file
    Public Sub Reset()

        TS_log.Flush()
        TS_log.Close()

    End Sub

    ' dgp rev 11/15/06 setup for application logging
    Public Sub Start_Logging()

        Dim test_path As String
        Dim valid_path As String
        If (fso.FolderExists(Environ("APPDATA"))) Then
            valid_path = Environ("APPDATA")
            test_path = fso.BuildPath(valid_path, "Flow Control")
            If (fso.FolderExists(test_path)) Then
                valid_path = test_path
            Else
                If (Create_Tree(test_path)) Then valid_path = test_path
            End If
        Else
            valid_path = CurDir()
        End If

        App_Log_File = fso.BuildPath(valid_path, Prefix + Format(Now(), "yyMMddhhmmss") + ".log")

        Try
            TS_log = New StreamWriter(App_Log_File, True)
            Logging_Flag = True
            Log_Info(Format(Now(), "MMM dd yyyy hh mm"))
        Catch ex As Exception

        End Try

    End Sub


    ' dgp rev 7/16/07 Log File Exists?
    Public Function Exists() As Boolean

        fso.FileExists(App_Log_File)

    End Function

    ' dgp rev 11/15/06 write logging info if logging is enabled
    Public Sub Log_Info(ByVal text As String)

        If (Logging_Flag) Then
            TS_log.WriteLine(text)
        End If

    End Sub

    ' dgp rev 11/13/06 clean up before exiting
    Public Sub Clean_Up()

        ' close the log file
        TS_log.Close()

    End Sub
    ' dgp rev 7/16/07 Create a new instance
    Public Sub New(ByVal name As String)

        Prefix = name
        Start_Logging()

    End Sub
End Class
