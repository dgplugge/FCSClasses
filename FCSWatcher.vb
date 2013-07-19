' Name:     FCS File Tracking Windows Service 
' Author:   Donald G Plugge
' Date:   3/19/09
' Purpose:  Class for tracking the creation of FCS files.
' 
Imports HelperClasses
Imports System.IO
Imports Microsoft.Win32
Imports FCS_Classes

Public Class FCSWatcher

    ' dgp rev 3/12/09 Create a File Watcher Service for FCS files
    Private WithEvents m_FileSystemWatcher As FileSystemWatcher
    Private mFileCount As Integer
    Private mFolderCount As Integer

    Private Shared mServiceName As String = "FCS Watcher"
    Public Shared ReadOnly Property ServiceName() As String
        Get
            Return mServiceName
        End Get
    End Property

    Public Enum FCSServiceCommands
        StartWatcher = 128
        StopWatcher
        RestartWatcher
        UpdatePath '
        WatcherStatus
    End Enum 'FCSServiceCommands

    ' dgp rev 3/19/09 
    Public ReadOnly Property XMLKeys() As Object
        Get
            If Dyn Is Nothing Then Return ""
            Return Dyn.GetKeys
        End Get
    End Property

    ' dgp rev 3/19/09 
    Public ReadOnly Property XMLFiles() As Object
        Get
            Return ValidFiles()
        End Get
    End Property

    ' dgp rev 3/19/09 
    Public ReadOnly Property XMLFolders() As Object
        Get
            Return ValidFolders()
        End Get
    End Property

    ' dgp rev 3/18/09
    Private Function ValidFiles() As ArrayList

        Dim Arr As New ArrayList
        If XMLKeys Is Nothing Then Return Arr

        Dim item
        For Each item In XMLKeys
            If item.ToString.ToUpper.Contains("FILE") Then
                If (System.IO.file.exists(Dyn.GetSetting(item))) Then Arr.Add(Dyn.GetSetting(item))
            End If
        Next
        Return Arr

    End Function

    ' dgp rev 3/18/09
    Private Function ValidFolders() As ArrayList

        Dim Arr As New ArrayList
        If XMLKeys Is Nothing Then Return Arr

        Dim item
        For Each item In XMLKeys
            If item.ToString.ToUpper.Contains("FOLDER") Then
                If (System.IO.Directory.exists(Dyn.GetSetting(item))) Then Arr.Add(Dyn.GetSetting(item))
            End If
        Next
        Return Arr

    End Function

    Private myLog As Diagnostics.EventLog

    ' dgp rev 1/5/2012
    Private Sub WriteLog(ByVal Txt As String)

        myLog = New Diagnostics.EventLog
        myLog.Source = "FCS Watcher"

        ' Write an informational entry to the event log.    
        myLog.WriteEntry(Txt)
        myLog.Close()

    End Sub

    Private Sub SetupEventLog()

        ' Add any initialization after the InitializeComponent() call.
        If Not Diagnostics.EventLog.SourceExists("FCS Watcher") Then
            ' Create the source, if it does not already exist.
            ' An event log source should not be created and immediately used.
            ' There is a latency time to enable the source, it should be created
            ' prior to executing the application that uses the source.
            ' Execute this sample a second time to use the new source.
            Diagnostics.EventLog.CreateEventSource("FCS Watcher", "Application")
            'The source is created.  Exit the application to allow it to be registered.
            Return
        End If

    End Sub

    Private mWatcherStatus As String = "No Status Yet"
    Public ReadOnly Property GetStatus As String
        Get
            Return mWatcherStatus
        End Get
    End Property

    ' dgp rev 6/29/2010 Retrieve the status of the currnet watcher
    Public Sub WatcherStatus()

        If (m_FileSystemWatcher Is Nothing) Then
            mWatcherStatus = "No Watcher"
        Else
            mWatcherStatus = "Watcher Status - " + m_FileSystemWatcher.Path
        End If

        If (mServiceMode) Then
            WriteLog(mWatcherStatus)
        Else
            Console.WriteLine(mWatcherStatus)
        End If

    End Sub

    ' dgp rev 6/29/2010 Log the last event
    Public Sub LogEvent()

        If mServiceMode Then
            WriteLog(mWatcherStatus)
        Else
            Console.WriteLine(mWatcherStatus)
        End If

    End Sub

    ' dgp rev 6/29/2010 Log the last event
    Public Sub LogEvent(ByVal txt As String)

        If mServiceMode Then
            WriteLog(txt)
        Else
            Console.WriteLine(txt)
        End If

    End Sub

    Private mCurWatcherPath As String = "F:\"


    ' dgp rev 6/29/2010 Function to run when client requests an update
    Public Function SyncWatchPath() As Boolean

        WriteLog("FCS Watcher Sync")
        SyncWatchPath = False
        Try
            If ValidRegPath("WatchPath") Then
                mCurWatcherPath = ReadRegPath("WatchPath")
                If (m_FileSystemWatcher IsNot Nothing) Then m_FileSystemWatcher.Path = mCurWatcherPath
                mWatcherStatus = "Path Changed to " + mCurWatcherPath
            End If
            SyncWatchPath = True
        Catch ex As Exception
            mWatcherStatus = "Error with path change"
        End Try

        LogEvent()

    End Function

    ' dgp rev 3/19/09 
    Private mOperPath = Nothing
    ' dgp rev 6/12/09 Operator's Path, where the XML database resides.
    Public ReadOnly Property OperPath() As String
        Get
            If mOperPath Is Nothing Then
                mOperPath = System.IO.Path.Combine(FlowStructure.FlowRoot, "Operator")
            End If
            Return mOperPath
        End Get
    End Property

    ' dgp rev 7/7/2010 Dynamic XML
    Private mDyn As HelperClasses.Dynamic = Nothing
    Public Property Dyn As HelperClasses.Dynamic
        Get
            If (HelperClasses.Utility.Create_Tree(OperPath)) Then
                If (CurXML = mDate) Then
                    If mDyn Is Nothing Then
                        mDyn = New HelperClasses.Dynamic(OperPath, CurXML + ".xml")
                    End If
                Else
                    mDyn = New HelperClasses.Dynamic(OperPath, CurXML + ".xml")
                    mFileCount = 0
                    mFolderCount = 0
                End If
            End If
            Return mDyn
        End Get
        Set(ByVal value As HelperClasses.Dynamic)

        End Set
    End Property

    Private mDate As String = Now.Date.ToLongDateString

    ' dgp rev 6/12/09 Current Tracking Date
    Public Property TrackingDate() As String
        Get
            Return mDate
        End Get
        Set(ByVal value As String)
            mDate = value
        End Set
    End Property

    ' dgp rev 7/7/2010 
    Public Function DisableWatcher() As Boolean

        If m_FileSystemWatcher Is Nothing Then Return False
        Try
            m_FileSystemWatcher.EnableRaisingEvents = False

        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    ' dgp rev 7/7/2010 
    Public Function EnableWatcher() As Boolean

        If m_FileSystemWatcher Is Nothing Then Return False
        Try
            m_FileSystemWatcher.EnableRaisingEvents = True

        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    ' dgp rev 7/7/2010 
    Public Function WatcherOn() As Boolean

        If m_FileSystemWatcher Is Nothing Then Return False
        Return m_FileSystemWatcher.EnableRaisingEvents

    End Function

    ' dgp rev 7/7/2010 
    Public Function WatcherExists() As Boolean

        Return m_FileSystemWatcher IsNot Nothing

    End Function

    ' dgp rev 7/7/2010 
    Public Function CreateWatcher() As Boolean

        Try
            m_FileSystemWatcher = New FileSystemWatcher
            m_FileSystemWatcher.Path = ReadRegPath("WatchPath")
            m_FileSystemWatcher.EnableRaisingEvents = True
            m_FileSystemWatcher.IncludeSubdirectories = True
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


    ' dgp rev 3/18/09 Initialize the Service
    Public Function StartService() As Boolean

        Try
            WriteLog("Starting FCS Watcher - " + Now.ToShortDateString())
        Catch ex As Exception
            WriteLog("FCS Watcher Error - " + ex.Message)
            Return False
        End Try
        Return True
    End Function

    ' dgp rev 3/18/09 Initialize the Service
    Public Function StopService() As Boolean

        WriteLog("Stopping FCS Watcher - " + Now.ToShortDateString())
        Try
            ' Add code here to perform any tear-down necessary to stop your service.
            Dyn.PutSetting("StopDate", Now.Date.ToLongDateString.ToString)
            Dyn.PutSetting("StopTime", Now.TimeOfDay.ToString)
            If m_FileSystemWatcher IsNot Nothing Then m_FileSystemWatcher.EnableRaisingEvents = False
        Catch ex As Exception
            WriteLog("FCS Watcher Error - " + ex.Message)
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 3/19/09 
    Private mKeyList() As String
    Public ReadOnly Property KeyList() As String()
        Get
            Return mKeyList
        End Get
    End Property

    ' dgp rev 3/19/09
    Private Function ValidateRegistry() As Boolean

        '        Dim RegPath As String = "System\CurrentControlSet\Services\" + Me.ProductName

        RegKey = Registry.LocalMachine.OpenSubKey(mRegPath, False)
        If (RegKey Is Nothing) Then Return False
        mKeyList = RegKey.GetValueNames
        If (KeyList.Length = 0) Then Return False
        Return (Array.IndexOf(KeyList, "WatchPath") >= 0)

    End Function

    ' dgp rev 7/7/2010 
    Public ReadOnly Property WatchPath As String
        Get
            Return ReadRegPath("WatchPath")
        End Get
    End Property

    Private RegKey As RegistryKey
    Private mRegPath As String = String.Format("System\CurrentControlSet\Services\{0}", mServiceName)
    Private mParPath = mRegPath + "\Parameters"

    ' dgp rev 7/7/2010 
    Public Function CreateDefReg() As Boolean

        Try
            RegKey = Registry.LocalMachine.CreateSubKey(mRegPath, RegistryKeyPermissionCheck.ReadWriteSubTree)
            RegKey.SetValue("WatchPath", mCurWatcherPath)
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 3/19/09 Establish the FlowRoot
    Private Function Establish_Config() As Boolean

        If Not WatcherExists() Then
            If (ValidRegPath("WatchPath")) Then
                Return CreateWatcher()
            Else
                If CreateDefReg() Then
                    Return CreateWatcher()
                Else
                    Return False
                End If
            End If
        Else
            If Not WatcherOn() Then EnableWatcher()
        End If
        Return True

    End Function

    Private mResults As String
    Private other As Integer
    Private tag

    ' dgp rev 3/18/09
    Private Sub ProcessEvent()

        If (IO.File.Exists(mResults)) Then
            mFileCount = mFileCount + 1
            tag = "File" + CStr(mFileCount)
            Dyn.PutSetting(tag, mResults)
        Else
            If (IO.Directory.Exists(mResults)) Then
                mFolderCount = mFolderCount + 1
                tag = "Folder" + CStr(mFolderCount)
                Dyn.PutSetting(tag, mResults)
            Else
                other = other + 1
                tag = "Other" + CStr(other)
            End If
        End If

    End Sub

    Private mSetupIndex As Integer

    ' dgp rev 3/18/09 Select Setup Parameter
    Public Function ValidRegPath(ByVal par As String) As Boolean

        '        Dim ParPath = RegPath + "\" + par
        '        Dim ParPath = mRegPath + "\Parameters"
        RegKey = Registry.LocalMachine.OpenSubKey(mRegPath, False)
        If (RegKey Is Nothing) Then Return False
        Try
            Dim fold = RegKey.GetValue(par)
            Return System.IO.Directory.Exists(fold)
        Catch ex As Exception
            Return False
        End Try

    End Function


    ' dgp rev 3/18/09 Select Setup Parameter
    Public Function ReadRegPath(ByVal par As String) As String

        '        Dim ParPath = RegPath + "\" + par
        '        Dim ParPath = mRegPath + "\Parameters"
        RegKey = Registry.LocalMachine.OpenSubKey(mRegPath, False)
        If (RegKey Is Nothing) Then Return ""
        Return RegKey.GetValue(par)

    End Function

    ' dgp rev 3/18/09 Select Setup Parameter
    Public Function ModifyReg(ByVal par As String, ByVal val As String) As Boolean

        ModifyReg = False
        RegKey = Registry.LocalMachine.OpenSubKey(mRegPath, True)
        If (Not RegKey Is Nothing) Then
            Try
                RegKey.SetValue(par, val)
                Return True
            Catch ex As Exception
            End Try
        End If

    End Function

    ' Watch for file creations.
    Private Sub m_FileSystemWatcher_Renamed(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles m_FileSystemWatcher.Renamed

        mResults = e.FullPath
        ProcessEvent()

    End Sub


    ' Watch for file creations.
    Private Sub m_FileSystemWatcher_Created(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles m_FileSystemWatcher.Created

        mResults = e.FullPath
        ProcessEvent()

    End Sub

    ' dgp rev 7/7/2010 
    Private mServiceMode As Boolean
    Public ReadOnly Property ServiceMode As Boolean
        Get
            Return mServiceMode
        End Get
    End Property

    ' dgp rev 7/7/2010 
    Public ReadOnly Property CurXML As String
        Get
            Return Now.Date.ToLongDateString
        End Get
    End Property

    ' dgp rev 4/8/09 New instance of service
    Public Sub New()

        mServiceMode = (Not Environment.UserInteractive)
        SetupEventLog()

        If (Me.Establish_Config) Then
            Dyn.PutSetting("StartDate", Now.Date.ToLongDateString.ToString)
            Dyn.PutSetting("StartTime", Now.TimeOfDay.ToString)
            If (ValidateRegistry()) Then
                mCurWatcherPath = ReadRegPath("WatchPath")
            Else
                If Not CreateDefReg() Then
                    WriteLog("FCS Watcher Registry Failure")
                End If
            End If
        Else
            WriteLog("FCS Watcher Configuration Failure")
        End If

    End Sub

    ' dgp rev 7/7/2010 
    Protected Overrides Sub Finalize()

        MyBase.Finalize()

    End Sub
End Class
