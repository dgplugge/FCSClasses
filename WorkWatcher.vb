' Name:     Work Tracking Class
' Author:   Donald G Plugge
' Date:   2/10/2011
' Purpose:  Class for tracking flow work
' 
Imports HelperClasses
Imports System.IO
Imports Microsoft.Win32
Imports FCS_Classes


Public Class WorkWatcher

    Private Shared WithEvents mWorkFileWatcher As FileSystemWatcher = Nothing
    Private Shared mWatcherStatus

    ' use delegates for Repeat
    Delegate Sub NewWorkEvent(ByVal FileName As String)
    Public Shared Event NewWorkEventHandler As NewWorkEvent

    ' dgp rev 2/10/2011 Retrieve the status of the currnet watcher
    Public Shared ReadOnly Property WatcherStatus() As String
        Get
            If (mWorkFileWatcher Is Nothing) Then
                mWatcherStatus = "No Watcher"
            Else
                mWatcherStatus = "Watcher Status - " + mWorkFileWatcher.Path
            End If
            Return mWatcherStatus
        End Get
    End Property

    ' dgp rev 2/10/2011
    Public Shared Function EnableWatcher() As Boolean

        If mWorkFileWatcher Is Nothing Then Return False
        Try
            mWorkFileWatcher.EnableRaisingEvents = True

        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    ' dgp rev 2/10/2011
    Public Shared Function WatcherOn() As Boolean

        If mWorkFileWatcher Is Nothing Then Return False
        Return mWorkFileWatcher.EnableRaisingEvents

    End Function

    ' dgp rev 2/10/2011 newly created file
    Private Shared mNewFile As String
    Public Shared ReadOnly Property Newfile As String
        Get
            Return mNewFile
        End Get
    End Property

    Private Shared mExportList As New Hashtable

    ' dgp rev 2/10/2011
    Private Shared Sub NewWorkFileHandler()

        If mExportList.Contains(mNewFile) Then
            If Not mExportList(mNewFile) = System.IO.File.GetLastWriteTime(mNewFile) Then
                mExportList(mNewFile) = System.IO.File.GetLastWriteTime(mNewFile)
                RaiseEvent NewWorkEventHandler(mNewFile)
            End If
        Else
            mExportList.Add(mNewFile, System.IO.File.GetLastWriteTime(mNewFile))
            RaiseEvent NewWorkEventHandler(mNewFile)
        End If

    End Sub

    ' dgp rev 2/10/2011
    Private Shared Sub NewWorkFileHandlerx()

        mTimeStamp = System.IO.File.GetLastWriteTime(mNewFile)
        If Not mPreviousTimeStamp = mTimeStamp Then
            mPreviousTimeStamp = mTimeStamp
            RaiseEvent NewWorkEventHandler(mNewFile)
        End If

    End Sub

    ' dgp rev 2/13/2012 TimeStamp of New Work
    Private Shared mTimeStamp As Date
    Public Shared ReadOnly Property TimeStamp As Date
        Get
            Return mTimeStamp
        End Get
    End Property
    Private Shared mPreviousTimeStamp As Date = Now()

    ' Watch for file creations.
    Private Shared Sub Work_Created(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles mWorkFileWatcher.Created

        mNewFile = e.FullPath
        NewWorkFileHandler()

    End Sub

    ' Watch for file creations.
    Private Shared Sub Work_Modified(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles mWorkFileWatcher.Changed

        mNewFile = e.FullPath
        NewWorkFileHandler()

    End Sub

    ' dgp rev 2/10/2011
    Public Shared Sub StartWatching(ByVal path)

        mWorkFileWatcher = New FileSystemWatcher(path)
        mWorkFileWatcher.IncludeSubdirectories = False
        EnableWatcher()

    End Sub
End Class
