' Author: Donald G Plugge
' Title: LogFiles
' Date: 11/24/2010
' Purpose: PV-Wave Log File Analysis 

Imports System.IO
Imports System.Text.RegularExpressions.Regex
Imports System.Text.RegularExpressions

Public Class LogFiles

    Private _logspec As String

    Sub New(ByVal logspec As String)
        ' TODO: Complete member initialization 
        _logspec = logspec
    End Sub

    Private Property logtext As String

    Function exists() As Boolean

        Return System.IO.File.Exists(_logspec)

    End Function

    Private sr As StreamReader
    Private RegExp As System.Text.RegularExpressions.Regex

    Private Function OpenLog() As Boolean

        Try
            sr = New StreamReader(_logspec)
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Private mCount = 0
    Public ReadOnly Property Count As Integer
        Get
            If mLines Is Nothing Then ReadLog()
            Return mLines.Count
        End Get
    End Property

    Private mLines As ArrayList

    Function ReadLog() As Boolean

        Dim done = False
        mLines = New ArrayList
        If Not OpenLog() Then Return False
        Do While (Not sr.EndOfStream)
            mLines.Add(sr.ReadLine)
        Loop

        Return mLines.Count > 0

    End Function

    Private mOrderedRoutines As ArrayList
    Private mOrderedTimes As ArrayList

    ' dgp rev 11/24/2010
    Public ReadOnly Property OrderedRoutines As ArrayList
        Get
            Return mOrderedRoutines
        End Get
    End Property

    ' dgp rev 11/24/2010
    Public ReadOnly Property OrderedTimes As ArrayList
        Get
            Return mOrderedTimes
        End Get
    End Property

    Private mMatchResults As ArrayList

    Public Function ExtractItems(ByVal key As String) As Integer

        Dim objMatchCollection As MatchCollection
        Dim objMatch As Match
        Dim item
        Dim match
        mMatchResults = New ArrayList
        ExtractAllTimings()
        For Each item In OrderedRoutines
            ' dgp rev 3/26/08 let's first remove the comment fields
            objMatchCollection = Regex.Matches(item, "^.*(" + key + ").*(\s\d+[:]\d+[:]\d+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            If objMatchCollection.Count > 0 Then
                For Each objMatch In objMatchCollection
                    If (objMatch.Groups.Count > 2) Then
                        match = objMatch.Groups.Item(2)
                        mMatchResults.Add(match.ToString)
                    End If
                Next
            End If
        Next
        Return mMatchResults.Count

    End Function

    Private Function ExtractItems(ByVal arr As ArrayList) As Integer

        Dim key As String
        Dim sum = 0
        For Each key In arr
            sum += ExtractItems(key)
        Next
        Return sum
    End Function


    ' dgp rev 11/24/2010 Extracting all timing information
    Function ExtractAllTimings() As Boolean

        mOrderedRoutines = New ArrayList
        Dim idx As Integer = 0
        If Count > 0 Then
            Do While (idx < Count)
                Try
                    If mLines.Item(idx).ToString.Length > 2 Then
                        If mLines.Item(idx).ToString.ToLower.Substring(0, 3).Contains("dgp") Then
                            mOrderedRoutines.Add(mLines.Item(idx).ToString)
                            idx += 1
                        End If
                    End If
                    idx += 1
                Catch ex As Exception
                    Return False
                End Try
            Loop
        End If
        Return True

    End Function

End Class
