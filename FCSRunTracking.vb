' Author: Donald G Plugge
' Date: 12/31/08
' Purpose: Tracking the last FCS run for a given user
Imports FCS_Classes
Imports HelperClasses

Public Class FCSRunTracking

    Private mUploadRoot As String = "\\Nt-eib-10-6b16\Upload"
    Public ReadOnly Property AriaRunRoot() As String
        Get
            Return system.io.path.combine(mUploadRoot, "FCSRun")
        End Get
    End Property

    Public Shared FTP_Path As String = "\\Nt-eib-10-6b16\FTP_root\runs"
    Public Shared RemoteSettingsPath As String = "\\Nt-eib-10-6b16\Upload\Settings"
    Public Shared RemoteMachinePath As String = "\\Nt-eib-10-6b16\Upload\Reserve\Machines"
    Public Shared RemoteUserPath As String = "\\Nt-eib-10-6b16\Upload\Reserve\Users"
    Private Shared objImp As New HelperClasses.RunAs_Impersonator
    Public PCtoVMS As Dynamic
    Public MachineInfo As Dynamic
    Public UserInfo As Dynamic
    Public RunLog As Dynamic
    Private mXMLRun As String
    Private mAriaRunList As ArrayList

    ' dgp rev 3/3/09 VMS events - NewVantagerun, VantageRuns, MAPUsers and VantageFiles
    Public Delegate Sub Event_NewVantageRunHand(ByVal info As Object)
    Public Shared Event Event_NewVantageRun As Event_NewVantageRunHand
    Public Delegate Sub Event_VantageRunsHand(ByVal info As Object)
    Public Shared Event Event_VantageRuns As Event_VantageRunsHand
    Public Delegate Sub Event_MAPUsersHand(ByVal info As Object)
    Public Shared Event Event_MAPUsers As Event_MAPUsersHand
    Public Delegate Sub Event_VantageFilesHand(ByVal info As Object)
    Public Shared Event Event_VantageFiles As Event_VantageFilesHand

    Private mVMSAccess As New VMSAccess

    Enum MachineTypes

        None = 0
        LSR_II = 1
        Aria = 2
        Vantage = 3

    End Enum

    ' dgp rev 2/27/09 Fill machine list
    Private Shared Sub FillList()

        mMachList = New ArrayList
        mMachList.Add("Select Machine")
        mMachList.Add("LSR_II")
        mMachList.Add("Aria")
        mMachList.Add("Vantage")

    End Sub

    ' dgp rev 2/27/09 Machine List
    Private Shared mMachList As ArrayList
    Public Shared ReadOnly Property MachList() As ArrayList
        Get
            If (mMachList Is Nothing) Then FillList()
            Return mMachList
        End Get
    End Property

    ' dgp rev 4/10/09 Machine Index
    Private mMachIndex As Integer
    Public Property MachIndex() As Integer
        Get
            Return mMachIndex
        End Get
        Set(ByVal value As Integer)
            mMachIndex = value
            mMachine = mMachList.Item(value)
            mAriaFileList = Nothing
            MachineInfo = New Dynamic(RemoteMachinePath, mMachine)
        End Set
    End Property

    Private mXMLAriaRun As String = ""
    Public Property XMLAriaRun() As String
        Get
            If (Not RunLog Is Nothing And RunLog.Exists(NCIUser)) Then
                mXMLAriaRun = RunLog.GetSetting(NCIUser)
            End If
            Return mXMLAriaRun
        End Get
        Set(ByVal value As String)
            If (Not RunLog Is Nothing) Then RunLog.PutSetting(NCIUser, value)
        End Set
    End Property

    Private mXMLVantageRun As String = ""
    Public Property XMLVantageRun() As String
        Get
            If (Not RunLog Is Nothing And RunLog.Exists(AssignedMap)) Then
                mXMLVantageRun = RunLog.GetSetting(AssignedMap)
            End If
            Return mXMLVantageRun
        End Get
        Set(ByVal value As String)
            If (Not RunLog Is Nothing And Not AssignedMap = "") Then RunLog.PutSetting(AssignedMap, value)
        End Set
    End Property

    ' dgp rev 1/14/08 Assigned Map
    ' dgp rev 1/14/08 Assigned Run
    ' dgp rev 1/14/08 Assigned NCIUser
    Enum RunStatus

        None = 0
        Held = 1
        Used = 2

    End Enum

    ' dgpr ev 7/16/09 
    Private mAriaRuns = Nothing
    Private mVantageRuns As ArrayList
    Private mRunsUsed As ArrayList
    Private mRunsHeld As ArrayList

    Private mCheckSum As String
    Private mFCSList As ArrayList

    ' dgp rev 1/14/09 Flags
    Private mDataFlag As Boolean
    Private mUserFlag As Boolean = False
    Private mMachineFlag As Boolean
    Private mRunFlag As Boolean

    Public RemoteInfo As Object
    ' dgp rev 10/16/08 
    Public Function CreateMap(ByVal user As String) As Boolean

        mImp.ImpersonateStart()
        PCtoVMS.PutSetting(NCIUser, user)
        CreateMap = PCtoVMS.Exists(NCIUser)
        mImp.ImpersonateStop()

    End Function

    Private mImp As HelperClasses.RunAs_Impersonator
    Public Property Impersonate() As HelperClasses.RunAs_Impersonator
        Get
            Return mImp
        End Get
        Set(ByVal value As HelperClasses.RunAs_Impersonator)
            mImp = value
        End Set
    End Property

    Private mValidMap As Boolean = False
    ' dgp rev 2/18/09 Clear the current user mapping
    Public Function Clear_Mapping() As Boolean

        mImp.ImpersonateStart()
        PCtoVMS.PutSetting(NCIUser, "")
        mImp.ImpersonateStop()
        mValidMap = False
        mAssignedMap = ""

    End Function

    ' dgp rev 2/18/09 Assigned Map
    Private mAssignedMap As String
    Public ReadOnly Property AssignedMap() As String
        Get
            If (Not mValidMap) Then Return ""
            Return mAssignedMap
        End Get
    End Property

    ' dgp rev 2/18/09 Lock NCI user to VMS user
    Public Function UnLockMap(ByVal user As String) As Boolean

        ' dgp rev 2/18/09 Already mapped
        If (Not mValidMap) Then Return False

        If (PCtoVMS Is Nothing) Then Return False

        mImp.ImpersonateStart()
        If (PCtoVMS.Exists(NCIUser)) Then
            PCtoVMS.RemoveSetting(NCIUser)
            mValidMap = False
            mAssignedMap = ""
        End If
        mImp.ImpersonateStop()


    End Function

    ' dgp rev 2/18/09 Lock NCI user to VMS user
    Public Function LockMap(ByVal user As String) As Boolean

        ' dgp rev 2/18/09 Already mapped
        If (mValidMap) Then Return False

        If (PCtoVMS Is Nothing) Then Return False
        mImp.ImpersonateStart()
        PCtoVMS.PutSetting(NCIUser, user)
        mImp.ImpersonateStop()
        mValidMap = True
        mAssignedMap = user

    End Function

    ' dgp rev 12/31/08 EIB NCIUser
    Private mUsername As String = System.Environment.GetEnvironmentVariable("username")
    Private mNCIUser = Nothing
    Public Property NCIUser() As String
        Get
            If mNCIUser Is Nothing Then mNCIUser = mUsername
            Return mNCIUser
        End Get
        Set(ByVal value As String)
            If mNCIUser IsNot Nothing Then
                If (mNCIUser.ToLower = value.ToLower) Then Exit Property
            End If
            mNCIUser = value
            mAriaRuns = Nothing
            mVantageRuns = Nothing
            mUserFlag = True
        End Set
    End Property

    ' dgp rev 2/13/09 Force Run into proper string format
    Public Shared Function RunFormat(ByVal run As Integer) As String

        RunFormat = (String.Format("R{0:D5}", run))

    End Function

    ' dgp rev 2/13/09 Force Run into proper string format
    Public Shared Function RunFormat(ByVal run As String) As String

        Dim num = CInt(run.ToUpper.Replace("R", ""))
        RunFormat = (String.Format("R{0:D5}", num))

    End Function

    ' dgp rev 2/13/09 Force Run into proper string format
    Public Shared Function RunNum(ByVal run As Integer) As Integer

        Return run

    End Function

    ' dgp rev 2/13/09 Force Run into proper string format
    Public Shared Function RunNum(ByVal run As String) As Integer

        Try
            Return CInt(run.ToUpper.Replace("R", ""))
        Catch ex As Exception
            Return 0
        End Try

    End Function

    ' dgp rev 2/13/09 Integer to String conversion
    Public Shared Function ConvertRun(ByVal run As Integer) As String

        ConvertRun = (String.Format("R{0:D5}", CInt(run)))

    End Function

    ' dgp rev 2/13/09 String to Integer conversion
    Public Shared Function ConvertRun(ByVal run As String) As Integer

        Try
            ConvertRun = CInt(run.ToUpper.Replace("R", ""))
        Catch ex As Exception
            Return 0
        End Try

    End Function

    ' dgp rev 12/1/08 NCIUser List Request complete
    Public Sub Reply_NewVantageRun(ByVal info As Object)

        RemoveHandler mVMSAccess.Event_NewVantageRun, AddressOf Reply_NewVantageRun
        RaiseEvent Event_NewVantageRun(info)
        RemoteInfo = info

    End Sub

    ' dgp rev 12/1/08 NCIUser List Request complete
    Public Sub Reply_MAPUsers(ByVal info As Object)

        RemoveHandler mVMSAccess.Event_MAPUsers, AddressOf Reply_MAPUsers
        RemoteInfo = info
        RaiseEvent Event_MAPUsers(info)

    End Sub

    Private mVantageFiles As ArrayList
    Private mCurVantageRun As String

    Public ReadOnly Property VantageFiles() As ArrayList
        Get
            Return mVantageFiles
        End Get
    End Property

    Public ReadOnly Property CurVantageRun() As Integer
        Get
            Return Me.mCurVantageRun
        End Get
    End Property

    ' dgp rev 12/1/08 NCIUser List Request complete
    Public Sub Reply_VantageFiles(ByVal info As Object)

        RemoveHandler mVMSAccess.Event_VantageFiles, AddressOf Reply_VantageFiles
        mVantageFiles = info
        RaiseEvent Event_VantageFiles(info)

    End Sub

    ' dgp rev 3/3/09 Vantage run list
    Public Property VantageRuns() As ArrayList
        Get
            If mVantageRuns Is Nothing Then Request_VantageRuns()
            Return mVantageRuns
        End Get
        Set(ByVal value As ArrayList)
            If (value Is Nothing) Then Return
            value.Sort()
            mVantageRuns = value
            LastVantageRun = ConvertRun(value.Item(value.Count - 1))
        End Set
    End Property

    ' dgp rev 3/3/09 List returned of Vantage Runs
    ' save list, flag and set last run
    Public Sub Reply_VantageRuns(ByVal info As Object)

        mVantageScanned = True
        RemoveHandler mVMSAccess.Event_VantageRuns, AddressOf Reply_VantageRuns
        VantageRuns = info
        RaiseEvent Event_VantageRuns(info)

    End Sub

    ' dgp rev 12/11/08 retrieve a NCIUser list from the VMS system
    Public Sub Request_MAPUsers()

        AddHandler mVMSAccess.Event_MAPUsers, AddressOf Reply_MAPUsers
        mVMSAccess.Request_MAPUsers()

    End Sub

    ' dgp rev 12/11/08 retrieve a NCIUser list from the VMS system
    Public Sub Request_NewVantageRun(ByVal run As String)

        AddHandler mVMSAccess.Event_NewVantageRun, AddressOf Reply_NewVantageRun
        mVMSAccess.Request_NewVantageRun(AssignedMap, run)

    End Sub

    Private mAriaFileList As ArrayList

    ' dgp rev 3/13/09 Scan for files in current run
    Private Sub ScanFiles()

        Select Case Me.MachIndex
            Case MachineTypes.Vantage
            Case MachineTypes.Aria
                mScanAriaRun(Me.GetMachineRun(True))
            Case MachineTypes.LSR_II
                mScanAriaRun(Me.GetMachineRun(True))
        End Select

    End Sub
    ' dgp rev 3/13/09 Enumerator
    Private Enum mRequestState
        err = -1
        immediate = 0
        wait = 1
    End Enum

    ' dgp rev 3/13/09 List of Files in current run
    Public ReadOnly Property AriaFileList() As ArrayList
        Get
            If (mAriaFileList Is Nothing) Then ScanFiles()
            Return mAriaFileList
        End Get
    End Property

    ' dgp rev 3/13/09 Scan Flow Server for files in current run
    Private Sub mScanAriaRun(ByVal run As String)

        mAriaFileList = New ArrayList
        Dim fullrun = ""

        Dim path As String = System.IO.Path.Combine(AriaRunRoot, NCIUser)
        If (Not System.IO.Directory.exists(path)) Then Exit Sub

        Dim item As Object
        If (System.IO.Directory.GetDirectories(path).Length = 0) Then Exit Sub
        'If (System.IO.Directory.GetDirectories(path).length > 0) Then
        For Each item In System.IO.Directory.GetDirectories(path)
            If (item.ToUpper.Contains(run.ToUpper)) Then
                fullrun = item.ToUpper
                Exit For
            End If
        Next
        If (fullrun = "") Then Exit Sub
        
        path = fullrun
        ' dgp rev 7/17/09 Determine the number of files in run
        Dim fil
        If (System.IO.Directory.GetFiles(path).Length > 0) Then
            For Each fil In System.IO.Directory.GetFiles(path)
                mAriaFileList.Add(System.IO.Path.GetFileName(fil))
            Next
            mAriaFileList.Sort()
        End If

    End Sub
    ' dgp rev 3/13/09 Get Aria File List from specific run number
    Private Sub mGetAriaFileList(ByVal num As Integer)

        mScanAriaRun(ConvertRun(num))

    End Sub

    ' dgp rev 12/11/08 retrieve a NCIUser list from the VMS system
    Private Sub mGetAriaFileList(ByVal run As String)

        mScanAriaRun(run)

    End Sub

    ' dgp rev 2/20/09 Wrapper around member 
    Public Sub GetAriaFileList(ByVal run As String)

        mImp.ImpersonateStart()
        mGetAriaFileList(run)
        mImp.ImpersonateStop()

    End Sub

    ' dgp rev 3/4/09 Remove an empty Aria run
    Public Sub RemoveAriaRun(ByVal run)

        Dim path As String = System.IO.Path.Combine(Me.AriaRunRoot, NCIUser)
        If (System.IO.Directory.Exists(path)) Then
            Dim item
            Dim info
            If (System.IO.Directory.GetDirectories(path).Length > 0) Then
                For Each item In System.IO.Directory.GetDirectories(path)
                    info = item.Split("_")
                    If (info(info.length - 1) = run) Then
                        If (System.IO.Directory.GetFiles(item).Length = 0) Then
                            Utility.DeleteTree(item)
                            mAriaRuns = Nothing
                        End If
                    End If
                Next
            End If
        End If

    End Sub

    ' dgp rev 7/16/09 reset the Aria runs so the info is rescanned
    Public Sub ResetAria()

        mAriaRuns = Nothing
        mReserved = Nothing

    End Sub
    ' dgp rev 7/16/09  
    Public ReadOnly Property AriaRuns() As ArrayList
        Get
            If (mAriaRuns Is Nothing) Then mAriaRuns = ScanAriaRuns()
            Return mAriaRuns
        End Get
    End Property

    Private mReserved = Nothing
    Public ReadOnly Property ReservedList() As ArrayList
        Get
            If mReserved Is Nothing Then ScanAriaRuns()
            Return mReserved
        End Get
    End Property

    ' dgp rev 12/11/08 retrieve a NCIUser list from the VMS system
    Private Function ScanAriaRuns() As ArrayList

        ScanAriaRuns = New ArrayList
        mReserved = New ArrayList

        Dim path As String = System.IO.Path.Combine(Me.AriaRunRoot, NCIUser)
        If (System.IO.Directory.Exists(path)) Then
            Dim info() As String
            Dim item
            Dim run
            If (System.IO.Directory.GetDirectories(path).Length > 0) Then
                For Each item In System.IO.Directory.GetDirectories(path)
                    info = System.IO.Path.GetFileNameWithoutExtension(item).ToString.Split("_")
                    If (item.ToUpper.Contains("RESERVED")) Then
                        mReserved.Add(info(info.Length - 1))
                    Else
                        run = info(info.Length - 1)
                        If Not run.length = 0 Then If (run.ToString.Substring(0, 1).ToUpper = "R") Then ScanAriaRuns.Add(run)
                    End If
                Next
                ' dgp rev 7/16/09 Make sure runs exist
                If ScanAriaRuns.Count = 0 Then Exit Function
                ScanAriaRuns.Sort()
                mLastAriaRun = ConvertRun(ScanAriaRuns.Item(ScanAriaRuns.Count - 1))
                GetAriaFileList(mLastAriaRun)
            End If
        End If

    End Function

    ' dgp rev 12/11/08 retrieve a NCIUser list from the VMS system
    Public Sub Request_VantageRuns()

        AddHandler mVMSAccess.Event_VantageRuns, AddressOf Reply_VantageRuns
        mVMSAccess.Request_VantageRuns(AssignedMap)

    End Sub

    ' dgp rev 12/11/08 retrieve a NCIUser list from the VMS system
    Public Sub Request_VantageFiles(ByVal run As String)

        Me.mCurVantageRun = run
        AddHandler mVMSAccess.Event_VantageFiles, AddressOf Reply_VantageFiles
        mVMSAccess.Request_VantageFiles(AssignedMap, run)

    End Sub

    ' dgp rev 1/13/09 VMS User
    Private mVMSUser As String
    Public Property VMSUser() As String
        Get
            Return mVMSUser
        End Get
        Set(ByVal value As String)
            mVMSUser = value
        End Set
    End Property

    ' dgp rev 1/14/08 Current Machine index
    Private mCurMachine As Int16

    ' dgp rev 1/13/09 VMS User
    Private mMachine As String
    Public Property Machine() As String
        Get
            Return mMachine
        End Get
        Set(ByVal value As String)
            mMachine = value
            mMachIndex = mMachList.IndexOf(value)
            mAriaFileList = Nothing
            MachineInfo = New Dynamic(RemoteMachinePath, mMachine)
        End Set
    End Property

    ' dgp rev 1/14/08 Current Assignments
    Private mAssignedData As String
    Private mAssignedRun As String
    Private mAssignedMachine As String
    Private mAssignedUser As String

    Public Property AssignedUser() As String
        Get
            Return mAssignedUser
        End Get
        Set(ByVal value As String)
            mAssignedUser = value
        End Set
    End Property

    ' dgp rev 12/31/08 Last VMS run number
    Private mVantageScanned As Boolean = False

    ' dgp rev 12/31/08 Last Logged Run
    Private mLastLoggedRun As Integer = 0
    Private mLogReadFlag As Boolean = False
    Public ReadOnly Property LastLoggedRun() As Integer
        Get
            If (Not mLogReadFlag) Then ReadRunLog()
            Return mLastLoggedRun
        End Get
    End Property

    ' dgp rev 2/20/09 Log the newly create run
    Public Function LogNewRun(ByVal run As String) As Boolean

        mImp.ImpersonateStart()
        LogNewRun = False
        Dim Run_File As String = System.IO.Path.Combine(FTP_Path, AssignedMap + ".vms")
        If (System.IO.File.Exists(Run_File)) Then
            Try
                Dim sw As New IO.StreamWriter(Run_File)
                sw.Write(ConvertRun(run))
                sw.Close()
                LogNewRun = True
            Catch ex As Exception
            End Try
        End If
        mImp.ImpersonateStop()

    End Function

    ' dgp rev 12/31/08 Last Run actually found on Vantage
    Private mLastVantageRun As Integer = 0
    Public Property LastVantageRun(ByVal strMode As Boolean) As Object
        Get
            If (Not mVantageScanned) Then Return 0
            ' dgp rev 3/3/09 if vantage scanned, then return the last run or 0
            If (strMode) Then
                Return RunFormat(mLastVantageRun)
            Else
                Return RunNum(mLastVantageRun)
            End If
        End Get
        Private Set(ByVal value As Object)
            If (value Is Nothing) Then Return
            mLastVantageRun = value
            Request_VantageFiles(mLastVantageRun)
        End Set
    End Property

    Public Property LastVantageRun() As Integer
        Get
            If (Not mVantageScanned) Then Return 0
            Return mLastVantageRun
        End Get
        Private Set(ByVal value As Integer)
            mLastVantageRun = value
        End Set
    End Property

    ' dgp rev 2/18/09 Create the Aria Run
    Public Function mCreateAriaRun(ByVal run As String) As Boolean

        run = ConvertRun(ConvertRun(run))

        Dim path As String = System.IO.Path.Combine(Me.AriaRunRoot, NCIUser)
        If (Not System.IO.Directory.Exists(path)) Then
            If (Not Utility.Create_Tree(path)) Then Return False
        End If

        Dim item
        For Each item In System.IO.Directory.GetDirectories(path)
            If (System.IO.Path.GetFileNameWithoutExtension(item).Contains(run)) Then Return False
        Next

        path = System.IO.Path.Combine(path, "Reserved_" + run)
        If System.IO.Directory.Exists(path) Then Return False

        Return Utility.Create_Tree(path)

    End Function

    ' dgp rev 3/3/09
    Public Function GetMachineRun(ByVal strmode As Boolean) As Object

        If (strmode) Then
            If (MachineInfo.Exists(NCIUser)) Then Return FCSRunTracking.RunFormat(MachineInfo.GetSetting(NCIUser))
            Return ""
        Else
            If (MachineInfo.Exists(NCIUser)) Then Return FCSRunTracking.RunNum(MachineInfo.GetSetting(NCIUser))
            Return 0
        End If

    End Function

    ' dgp rev 3/12/09 Check VMS System
    Private Function CheckVMS(ByVal user As String, ByVal run As String) As Boolean

    End Function

    ' dgp rev 3/12/09 Check Flow Server
    Private Function CheckServer(ByVal user As String, ByVal run As String) As Boolean

        CheckServer = True
        Dim path As String = System.IO.Path.Combine(Me.AriaRunRoot, NCIUser)
        If (System.IO.Directory.Exists(path)) Then
            Dim item
            For Each item In System.IO.Directory.GetDirectories(path)
                If (System.IO.Path.GetFileNameWithoutExtension(item).ToUpper.Contains(run)) Then
                    If (System.IO.Directory.GetFiles(path).Length = 0) Then Return False
                End If
            Next
        End If

    End Function

    ' dgp rev 3/13/09 Does the current machine run have an empty run folder
    Public Function EmptyMachineRun() As Boolean



    End Function

    ' dgp rev 3/3/09 Check Machine
    Public Function MachineRunExists() As Boolean

        MachineRunExists = False
        If (Not MachineInfo.Exists(NCIUser)) Then Exit Function

        Dim run = MachineInfo.GetSetting(NCIUser)
        run = ConvertRun(ConvertRun(run))

        If (CheckServer(NCIUser, run)) Then MachineInfo.RemoveSetting(NCIUser)

    End Function

    ' dgp rev 3/2/09 Log the new Run information
    Private Sub mLogNewRun(ByVal run As String)

        run = ConvertRun(ConvertRun(run))

        UserInfo.PutSetting(Machine, run)
        MachineInfo.PutSetting(NCIUser, run)

    End Sub

    ' dgp rev 2/18/09 Create the Aria Run
    Public Function CreateAriaRun(ByVal run As String) As Boolean

        mImp.ImpersonateStart()
        mCreateAriaRun(run)
        mLogNewRun(run)
        mImp.ImpersonateStop()

    End Function

    ' dgp rev 12/31/08 Last Aria run number
    Private mLastAriaRun = Nothing
    Public ReadOnly Property LastAriaRun() As Integer
        Get
            If mLastAriaRun Is Nothing Then ScanAriaRuns()
            Return mLastAriaRun
        End Get
    End Property

    Public ReadOnly Property LastAriaRun(ByVal strMode As Boolean) As Object
        Get
            If (strMode) Then
                Return RunFormat(LastAriaRun())
            Else
                Return RunNum(LastAriaRun())
            End If
        End Get
    End Property

    ' dgp rev 12/31/08 Last run number
    Public ReadOnly Property LastRun() As Integer
        Get
            If (Me.LastVantageRun > Me.LastAriaRun) Then Return LastVantageRun
            Return LastAriaRun
        End Get
    End Property

    ' dgp rev 12/31/08 Last run number 
    ' notice: must return object, as it can be integer or string
    Public ReadOnly Property LastRun(ByVal strmode As Boolean) As Object
        Get
            If (strmode) Then
                If (Me.LastVantageRun > Me.LastAriaRun) Then Return RunFormat(mLastVantageRun)
                Return RunFormat(mLastAriaRun)
            Else
                If (Me.LastVantageRun > Me.LastAriaRun) Then Return RunNum(LastVantageRun)
                Return RunNum(LastAriaRun)
            End If
        End Get
    End Property

    ' dgp rev 11/1/07 Read File
    ' dgp rev 11/1/07 Last Run is stored in VMS named text file
    ' cross referenced with PC name via XML file 
    Private Sub mReadRunLog()

        mLogReadFlag = True
        mLastLoggedRun = 0
        Dim Run_File As String = system.io.path.combine(FTP_Path, AssignedMap + ".vms")
        If (System.IO.file.exists(Run_File)) Then
            Dim sr As New IO.StreamReader(Run_File)
            Dim raw As String = sr.ReadToEnd
            sr.Close()
            Try
                mLastLoggedRun = CInt(raw)
            Catch ex As Exception
            End Try
        End If

    End Sub

    ' dgp rev 2/20/09
    Private Sub ReadRunLog()

        mImp.ImpersonateStart()
        mReadRunLog()
        mImp.ImpersonateStop()

    End Sub

    Private mMapChecked As Boolean = False

    ' dgp rev 2/18/09 Does user have a valid mapping
    Public ReadOnly Property IsMapped() As Boolean
        Get
            Return mValidMap
        End Get
    End Property

    ' dgp rev 2/13/09 Check for NCI to VMS mapping
    Public Sub ReadMapFile()

        mMapChecked = True
        mImp.ImpersonateStart()
        If (PCtoVMS.Exists(NCIUser)) Then
            mAssignedMap = PCtoVMS.GetSetting(NCIUser)
            mValidMap = True
        End If
        mImp.ImpersonateStop()

    End Sub

    ' dgp rev 3/6/09 Initialize the Tracker
    Private Sub Init(ByVal username)

        NCIUser = username

        PCtoVMS = New Dynamic(RemoteSettingsPath, "PCtoVMS")
        RunLog = New Dynamic(RemoteSettingsPath, "RunLog")

        UserInfo = New Dynamic(RemoteUserPath, username)
        ReadMapFile()

    End Sub

    ' dgp rev 2/18/09 New FCS Run Tracking object
    Public Sub New(ByVal Username As String, ByVal objImp As HelperClasses.RunAs_Impersonator)

        mImp = objImp
        mImp.ImpersonateStart()
        Init(Username)
        mImp.ImpersonateStop()

    End Sub
    ' dgp rev 2/18/09 New FCS Run Tracking object
    Public Sub New(ByVal Username As String)

        mImp = New RunAs_Impersonator
        Init(Username)

    End Sub
End Class
