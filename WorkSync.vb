Imports System.IO

' Name:     Sync Work
' Author:   Donald G Plugge
' Date:     9/21/2012
' Purpose:  Class to handle work synchronization

Public Class WorkSync

    Public Delegate Sub TransferBackupHandler(ByVal SomeString As String)
    Public Delegate Sub TransferErrorHandler(ByVal SomeString As String)
    Public Shared Event TransferTargetEvent As TransferBackupHandler
    Public Shared Event TransferErrorEvent As TransferBackupHandler
    Public Shared Event TransferDoneEvent As TransferBackupHandler

    Private Shared objBackupThread As Threading.Thread

    Public Shared Sub LaunchServerVerifyThread()

        objBackupThread = New Threading.Thread(New Threading.ThreadStart(AddressOf BackupMissingServerData))
        objBackupThread.Name = "StartVerify"

        ' dgp rev 10/19/09 Reinsert the event handler
        AddHandler TransferTargetEvent, AddressOf BackupEvent
        AddHandler TransferDoneEvent, AddressOf BackupDone

        objBackupThread.Start()

    End Sub



    ' 
    Public Shared Sub LaunchVerifyThread()

        objBackupThread = New Threading.Thread(New Threading.ThreadStart(AddressOf StartVerify))
        objBackupThread.Name = "StartVerify"

        ' dgp rev 10/19/09 Reinsert the event handler
        AddHandler TransferTargetEvent, AddressOf BackupEvent
        AddHandler TransferDoneEvent, AddressOf BackupDone

        objBackupThread.Start()

    End Sub

    Public Shared Sub LaunchBackupThread()

        objBackupThread = New Threading.Thread(New Threading.ThreadStart(AddressOf StartBackup))
        objBackupThread.Name = "StartBackup"

        ' dgp rev 10/19/09 Reinsert the event handler
        AddHandler TransferTargetEvent, AddressOf BackupEvent
        AddHandler TransferDoneEvent, AddressOf BackupDone

        objBackupThread.Start()

    End Sub

    Public Shared Sub LaunchRestoreThread()

        objBackupThread = New Threading.Thread(New Threading.ThreadStart(AddressOf StartRestore))
        objBackupThread.Name = "StartRestore"

        ' dgp rev 10/19/09 Reinsert the event handler
        AddHandler TransferTargetEvent, AddressOf BackupEvent
        AddHandler TransferDoneEvent, AddressOf BackupDone

        objBackupThread.Start()

    End Sub


    Private Shared mCurUser = Environment.UserName
    Public Shared Property CurUser As String
        Get
            Return mCurUser
        End Get
        Set(value As String)
            mCurUser = value
        End Set
    End Property

    Private Shared mServer As String = "NCI-01855598"
    Private Shared mShare As String = "FlowRoot"
    Public Shared ReadOnly Property Server As String
        Get
            Return mServer
        End Get
    End Property
    Public Shared ReadOnly Property Share As String
        Get
            Return mShare
        End Get
    End Property


    Private Shared mUserLocalWorkRoot = Nothing
    Public Shared ReadOnly Property UserLocalWorkRoot As String
        Get
            If mUserLocalWorkRoot Is Nothing Then mUserLocalWorkRoot = System.IO.Path.Combine(UserLocalPath, "Work")
            Return mUserLocalWorkRoot
        End Get
    End Property

    Private Shared mUserRemoteWorkRoot = Nothing
    Public Shared ReadOnly Property UserRemoteWorkRoot As String
        Get
            If mUserRemoteWorkRoot Is Nothing Then mUserRemoteWorkRoot = System.IO.Path.Combine(UserRemotePath, "Work")
            Return mUserRemoteWorkRoot
        End Get
    End Property

    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UserRemoteWorkRootExists As Boolean
        Get
            Return System.IO.Directory.Exists(UserRemoteWorkRoot)
        End Get
    End Property

    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UserLocalWorkRootExists As Boolean
        Get
            Return System.IO.Directory.Exists(UserLocalWorkRoot)
        End Get
    End Property

    Private Shared mUserRemotePath = Nothing
    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UserRemotePath As String
        Get
            If mUserRemotePath Is Nothing Then mUserRemotePath = System.IO.Path.Combine(UsersRemoteRoot, CurUser)
            Return mUserRemotePath
        End Get
    End Property

    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UsersRemotePathExists As Boolean
        Get
            Return System.IO.Directory.Exists(UserRemotePath)
        End Get
    End Property



    Private Shared mUsersRemoteRoot = Nothing
    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UsersRemoteRoot As String
        Get
            If mUsersRemoteRoot Is Nothing Then mUsersRemoteRoot = System.IO.Path.Combine(String.Format("\\{0}\{1}", Server, Share), "Users")
            Return mUsersRemoteRoot
        End Get
    End Property

    Private Shared mUserLocalPath = Nothing
    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UserLocalPath As String
        Get
            If mUserLocalPath Is Nothing Then mUserLocalPath = System.IO.Path.Combine(UsersLocalRoot, CurUser)
            Return mUserLocalPath
        End Get
    End Property

    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UserLocalPathExists As Boolean
        Get
            Return System.IO.Directory.Exists(UserLocalPath)
        End Get
    End Property


    ' dgp rev 9/24/2012 
    Private Shared mUsersLocalRoot = Nothing
    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UsersLocalRoot As String
        Get
            If mUsersLocalRoot Is Nothing Then mUsersLocalRoot = System.IO.Path.Combine(FlowStructure.FlowRoot, "Users")
            Return mUsersLocalRoot
        End Get
    End Property

    ' dgp rev 9/24/2012 
    Public Shared ReadOnly Property UsersRemoteRootExists As Boolean
        Get
            Return System.IO.Directory.Exists(UsersRemoteRoot)
        End Get
    End Property


    Private Shared mUsersLocalList = Nothing
    Public Shared ReadOnly Property UsersLocalList As ArrayList
        Get
            Dim USER
            If mUsersLocalList Is Nothing Then
                mUsersLocalList = New ArrayList
                If System.IO.Directory.Exists(UsersLocalRoot) Then
                    If Not System.IO.Directory.GetDirectories(UsersLocalRoot).Count = 0 Then
                        For Each user In System.IO.Directory.GetDirectories(UsersLocalRoot)
                            Try
                                If NIHNet.FlowLabUsers.Contains(System.IO.Path.GetFileNameWithoutExtension(USER).ToLower) Then
                                    mUsersLocalList.add(System.IO.Path.GetFileNameWithoutExtension(USER))
                                End If
                            Catch ex As Exception
                                If System.IO.Path.GetFileNameWithoutExtension(USER).ToLower = CurUser.ToLower Then
                                    mUsersLocalList.add(System.IO.Path.GetFileNameWithoutExtension(USER))
                                End If
                            End Try
                        Next
                    End If
                End If
            End If
            If mUsersLocalList.Count = 0 Then mUsersLocalList.add(CurUser)
            Return mUsersLocalList
        End Get
    End Property

    Public Shared ReadOnly Property ValidRemoteWork As Boolean
        Get
            If UserRemoteWorkRootExists Then
                Return Not System.IO.Directory.GetDirectories(UserRemoteWorkRoot).Count = 0
            End If
            Return False
        End Get
    End Property

    Public Shared ReadOnly Property ProjectsRemoteNames As ArrayList
        Get
            Dim arr = New ArrayList
            Dim proj
            If ValidRemoteWork Then
                For Each proj In System.IO.Directory.GetDirectories(UserRemoteWorkRoot)
                    arr.Add(System.IO.Path.GetFileNameWithoutExtension(proj))
                Next
            End If
            Return arr
        End Get
    End Property

    Public Shared ReadOnly Property ProjectsRemoteList As ArrayList
        Get
            Dim proj
            Dim arr = New ArrayList
            If ValidRemoteWork Then
                For Each proj In System.IO.Directory.GetDirectories(UserRemoteWorkRoot)
                    arr.Add(proj)
                Next
            End If
            Return arr
        End Get
    End Property

    Private Shared Sub DirectoryCopy( _
            ByVal sourceDirName As String, _
            ByVal destDirName As String, _
            ByVal copySubDirs As Boolean)

        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        ' If the source directory does not exist, throw an exception.
        If Not System.IO.Directory.Exists(sourceDirName) Then
            Throw New DirectoryNotFoundException( _
                "Source directory does not exist or could not be found: " _
                + sourceDirName)
        End If

        ' If the destination directory does not exist, create it.
        If Not Directory.Exists(destDirName) Then
            Directory.CreateDirectory(destDirName)
        End If

        Dim file
        For Each file In System.IO.Directory.GetFiles(sourceDirName)

            ' Create the path to the new copy of the file.
            Dim temppath As String = Path.Combine(destDirName, System.IO.Path.GetFileName(file))

            ' Copy the file.
            Try
                If Not System.IO.File.Exists(temppath) Then
                    System.IO.File.Copy(file, temppath, False)
                    If System.IO.Path.GetFileName(temppath).ToLower = "fcs_files.lis" Then mFCS_File_list.Add(temppath)
                    RaiseEvent TransferTargetEvent(temppath)
                End If
            Catch ex As Exception
                RaiseEvent TransferErrorEvent(String.Format("{0} {1}", ex.Message, temppath))
            End Try
        Next file

        Dim subdir
        ' If copySubDirs is true, copy the subdirectories.
        If copySubDirs Then

            For Each subdir In System.IO.Directory.GetDirectories(sourceDirName)

                ' Create the subdirectory.
                Dim temppath As String = _
                    Path.Combine(destDirName, System.IO.Path.GetFileName(subdir))

                ' Copy the subdirectories.
                DirectoryCopy(subdir, temppath, copySubDirs)
            Next subdir
        End If
    End Sub


    Private Shared mFCS_File_list As ArrayList
    Private Shared mDataRestoreList = Nothing
    Public Shared ReadOnly Property DataRestoreList As Hashtable
        Get
            If mDataRestoreList Is Nothing Then CreateRestoreList()
            Return mDataRestoreList
        End Get
    End Property


    ' dgp rev 
    Private Shared Sub StartRestore()

        mFCS_File_list = New ArrayList
        If UserLocalWorkRootExists Then
            DirectoryCopy(UserRemoteWorkRoot, UserLocalWorkRoot, True)
        Else
            If HelperClasses.Utility.Create_Tree(UserLocalWorkRoot) Then
                DirectoryCopy(UserRemoteWorkRoot, UserLocalWorkRoot, True)
            End If
        End If

        Dim item As DictionaryEntry
        For Each item In DataRestoreList
            DirectoryCopy(item.Value, System.IO.Path.Combine(FlowStructure.Data_Root, item.Key), True)
        Next

        RaiseEvent TransferDoneEvent("Done")

    End Sub

    Private Shared Sub BackupWork()

        If UserRemoteWorkRootExists Then
            DirectoryCopy(UserLocalWorkRoot, UserRemoteWorkRoot, True)
        Else
            If HelperClasses.Utility.Create_Tree(UserRemoteWorkRoot) Then
                DirectoryCopy(UserLocalWorkRoot, UserRemoteWorkRoot, True)
            End If
        End If

    End Sub

    Private Shared Function VerifyData() As Boolean



    End Function

    ' dgp rev 
    Private Shared Sub StartBackup()

        BackupWork()
        If VerifyData() Then

        End If
        RaiseEvent TransferDoneEvent("Done")

    End Sub

    Private Shared Sub BackupDone(SomeString As String)


    End Sub

    Private Shared Sub BackupEvent(SomeString As String)


    End Sub


    Public Shared ReadOnly Property ProblemLists As ArrayList
        Get
            If mLocalFCSFileList Is Nothing Then ScanMissingData()
            Return mLocalFCSFileList
        End Get
    End Property

    Private Shared Sub ScanMissingData()

        Dim fcslst As FCS_List
        Dim proj
        Dim sess
        mLocalFCSFileList = New ArrayList
        If UserLocalWorkRootExists Then
            For Each proj In System.IO.Directory.GetDirectories(UserLocalWorkRoot)
                For Each sess In System.IO.Directory.GetDirectories(proj)
                    fcslst = New FCS_List(sess)
                    If Not fcslst.AllValid Then
                        mLocalFCSFileList.Add(sess)
                    End If
                Next
            Next
        End If

    End Sub

    Private Shared mReportText As String = ""
    Private Shared ReadOnly Property ReportText As String
        Get
            Return mReportText
        End Get
    End Property


    Private Shared Sub CreateRestoreList()

        Dim lis
        For Each Lis In ProblemLists

            Dim RunName = ""
            Dim data_list = New FCS_List(Lis)
            Dim Unique_Runs As New ArrayList
            mDataRestoreList = New Hashtable
            ' dgp rev 9/25/2012 does data exists locally?
            If Not data_list.AllValid Then
                RunName = data_list.ListedRuns(0)
                If Not Unique_Runs.Contains(RunName) Then
                    Unique_Runs.Add(RunName)
                    If FlowServer.ServerUserRuns(CurUser).Contains(RunName) Then
                        mDataRestoreList.Add(RunName, FlowServer.ServerUserRunPath(CurUser, RunName))
                    Else
                        If FlowServer.ServerDataFindRun(RunName).Count = 0 Then
                            mReportText = mReportText + vbCrLf + String.Format("Run {0} not found", RunName)
                        Else
                            If Not mDataRestoreList.ContainsKey(RunName) Then
                                mDataRestoreList.Add(RunName, FlowServer.ServerDataFindRun(RunName).Item(0))
                            End If
                        End If
                    End If
                End If
            End If
        Next

    End Sub

    Private Shared mAllLocalFileLists = Nothing
    Public Shared ReadOnly Property AllLocalFileLists As ArrayList
        Get
            If mAllLocalFileLists Is Nothing Then ScanAllLocalFileLists()
            Return mAllLocalFileLists
        End Get
    End Property

    Private Shared mAllLocalRuns = Nothing
    Public Shared ReadOnly Property AllLocalRuns As ArrayList
        Get
            If mAllLocalRuns Is Nothing Then ScanAllLocalRuns()
            Return mAllLocalRuns
        End Get
    End Property

    Private Shared mMissingServerRuns = Nothing
    Public Shared ReadOnly Property MissingServerRuns
        Get
            If mMissingServerRuns Is Nothing Then ScanServerForRuns()
            Return mMissingServerRuns
        End Get
    End Property

    Private Shared mExtraDataList = Nothing
    Private Shared Sub ScanServerForRuns()

        mMissingServerRuns = New ArrayList
        Dim RunName
        For Each RunName In AllLocalRuns
            If Not FlowServer.ServerUserRuns(CurUser).Contains(RunName) Then
                If FlowServer.ServerDataFindRun(RunName).Count = 0 Then
                    mMissingServerRuns.Add(RunName)
                End If
            End If
        Next

    End Sub

    Private Shared Sub AllLocalDataToServer()

        mExtraDataList = New ArrayList
        Dim RunPath
        Dim RunName
        For Each RunPath In System.IO.Directory.GetDirectories(FlowStructure.Data_Root)
            RunName = System.IO.Path.GetFileNameWithoutExtension(RunPath)
            If Not FlowServer.ServerUserRuns(CurUser).Contains(RunName) Then
                If FlowServer.ServerDataFindRun(RunName).Count = 0 Then
                    mExtraDataList.Add(RunName)
                End If
            End If
        Next

        Dim target
        Dim source
        Dim run
        For Each run In mExtraDataList
            target = System.IO.Path.Combine(System.IO.Path.Combine(String.Format("\\{0}\{1}", Server, Share), "Data"), run)
            source = System.IO.Path.Combine(FlowStructure.Data_Root, run)
            If System.IO.Directory.Exists(source) Then
                If Not System.IO.Directory.GetFiles(source).Count = 0 Then
                    If Not System.IO.Directory.Exists(target) Then
                        DirectoryCopy(source, target, False)
                    End If
                End If
            End If
        Next

    End Sub

    Public Shared Sub BackupMissingServerData()

        Dim source
        Dim target
        Dim run
        Try
            For Each run In MissingServerRuns
                target = System.IO.Path.Combine(System.IO.Path.Combine(UserRemotePath, "Data"), run)
                source = System.IO.Path.Combine(FlowStructure.Data_Root, run)
                If System.IO.Directory.Exists(source) Then DirectoryCopy(source, target, False)
            Next
            AllLocalDataToServer()

        Catch ex As Exception

        End Try
        RaiseEvent TransferDoneEvent("Done")

    End Sub

    Private Shared Sub ScanAllLocalRuns()

        Dim RunName = ""
        Dim data_list As FCS_List
        mAllLocalRuns = New ArrayList
        Dim lis
        For Each lis In AllLocalFileLists
            data_list = New FCS_List(System.IO.Path.GetDirectoryName(lis))
            ' dgp rev 9/25/2012 does data exists locally?
            RunName = data_list.ListedRuns(0)
            If Not mAllLocalRuns.Contains(RunName) Then
                mAllLocalRuns.Add(RunName)
            End If
        Next

    End Sub

    Private Shared Sub ScanAllLocalFileLists()

        Dim fcslst As FCS_List
        mAllLocalFileLists = New ArrayList
        If UserLocalWorkRootExists Then
            Dim proj
            Dim sess
            For Each proj In System.IO.Directory.GetDirectories(UserLocalWorkRoot)
                For Each sess In System.IO.Directory.GetDirectories(proj)
                    fcslst = New FCS_List(sess)
                    mAllLocalFileLists.Add(fcslst.List_Spec)
                Next
            Next
        End If

    End Sub

    Private Shared mLocalFCSFileList = Nothing
    Private Shared Sub ScanFCSList()

        ScanMissingData()

        Dim item As DictionaryEntry
        For Each item In DataRestoreList
            DirectoryCopy(item.Value, System.IO.Path.Combine(FlowStructure.Data_Root, item.Key), True)
        Next

    End Sub

    Private Shared Sub StartVerify()

        If UserLocalWorkRootExists Then

            ScanFCSList()

        End If
        RaiseEvent TransferDoneEvent("Done")


    End Sub


End Class
