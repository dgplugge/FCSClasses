' Name:     FCS Server Download Routines
' Author:   Donald G Plugge
' Date:     1/26/07
' Purpose:  Class to handle operator downloads to the server
Imports HelperClasses
Imports FCS_Classes
Imports System.Management

Public Class FCSUpload

    Public Enum Stage
        Invalid = -1
        Mismatch = -2
        [Nothing] = 0
        Sync = 1
        Ready = 2
        Uploaded = 3
        Stored = 4
    End Enum

    ' use delegates for Repeat
    Delegate Sub delRepeat(ByVal Xname As String)
    Public Event evtRepeat As delRepeat

    Public Shared objImp As New RunAs_Impersonator
    Public objFCS_File As FCS_Classes.FCS_File
    Public objVMS As VMSAccess

    Public Target_Path As String

    Public Upload_Root As String
    Public DataServer As String = "nt-eib-10-6b16"
    Public VMSFlag As Boolean = False

    Public Shared ServerAccount As String
    Public Shared ServerPassword As String
    Public wmiOptions As New ConnectionOptions

    Public ServerOn As Boolean = False

    Public Dyn As New Dynamic("FCSServer")
    Public PCtoVMS As Dynamic

    Public FlagServer As Boolean = False
    Public FlagLock As Boolean = False

    Public Last_Drive As String = ""
    Public Flow_Root As String = ""

    Public Shared FTP_Path As String = "\\Nt-eib-10-6b16\FTP_root\runs"
    Public Shared AriaRunRoot As String = "\\Nt-eib-10-6b16\Upload\FCSRun"
    Public Shared RemoteSettingsPath As String = "\\Nt-eib-10-6b16\Upload\Settings"
    Public Shared RemoteMachinePath As String = "\\Nt-eib-10-6b16\Upload\Reserve\Machines"
    Public Shared RemoteUserPath As String = "\\Nt-eib-10-6b16\Upload\Reserve\Users"
    Public Last_Run_Str As String
    Public Last_Run_Num As Int16

    Private mMatchList As ArrayList
    Private mReserveList As ArrayList

    ' dgp rev 3/11/09 Valid Store Root
    Private mValidStore As Boolean
    Public ReadOnly Property ValidStore() As Boolean
        Get
            Return mValidStore
        End Get
    End Property

    ' dgp rev 3/11/09 Valid Store Root
    Private mStoreRoot As String
    Public Property StoreRoot() As String
        Get
            Return mStoreRoot
        End Get
        Set(ByVal value As String)
            If (System.IO.Directory.exists(value)) Then
                mStoreRoot = value
                mValidStore = True
            End If
        End Set
    End Property

    Private mStoredFlag As Boolean = False
    Public ReadOnly Property StoredFlag() As Boolean
        Get
            Return mStoredFlag
        End Get
    End Property

    ' dgp rev 3/11/09 Valid Run Store
    Public ReadOnly Property StoreRunPath() As String
        Get
            If (ValidStore) Then Return system.io.path.combine(StoreRoot, "FCSRun")
            Return ""
        End Get
    End Property

    ' dgp rev 3/11/09 Valid Run Store
    Public ReadOnly Property StoreExperPath() As String
        Get
            If (ValidStore) Then Return system.io.path.combine(StoreRoot, "FCSExper")
            Return ""
        End Get
    End Property

    Private mOrigPath As String
    Public ReadOnly Property OrigPath() As String
        Get
            Return mOrigPath
        End Get
    End Property

    ' dgp rev 8/1/08
    Public Function Store_Files() As Boolean

        Store_Files = False
        If (System.IO.Directory.exists(OrigPath)) Then
            If (System.IO.Directory.exists(StoreRoot)) Then
                Dim path = system.io.path.combine(StoreRoot, Format(Now(), "yyyyMMdd"))
                path = system.io.path.combine(path, Me.AssignedUser)
                If Utility.Create_Tree(path) Then
                    path = system.io.path.combine(path, Full_Name)
                    Try
                        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                        System.IO.File.Copy(OrigPath, path)
                        Store_Files = True
                    Catch ex As Exception

                    End Try
                    Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                End If
            End If
        End If

    End Function

    ' dgp rev 8/1/08
    Public Function Store_Exper() As Boolean

        Store_Exper = False
        If (System.IO.Directory.exists(OrigPath)) Then
            If (System.IO.Directory.exists(StoreExperPath)) Then
                Dim path = system.io.path.combine(StoreExperPath, Format(Now(), "yyyyMMdd"))
                path = system.io.path.combine(path, Me.AssignedUser)
                If Utility.Create_Tree(path) Then
                    path = system.io.path.combine(path, Full_Name)
                    Try
                        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                        System.IO.File.Move(OrigPath, path)
                        Store_Exper = True
                    Catch ex As Exception

                    End Try
                    Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                End If
            End If
        End If

    End Function

    ' dgp rev 3/6/09 Reserve List
    Public ReadOnly Property ReserveList() As ArrayList
        Get
            UserReserved()
            Return mReserveList
        End Get
    End Property
    ' dgp rev 3/6/09 Scan for any run matches
    Private Sub UserReserved()

        If (Upload_Root Is Nothing) Then
            Target_Path = AriaRunRoot
        Else
            Target_Path = String.Format("\\{0}{1}", DataServer, Upload_Root)
        End If
        Target_Path = system.io.path.combine(Target_Path, AssignedUser)

        mReserveList = New ArrayList
        Dim info
        Dim item
        If (System.IO.Directory.exists(Target_Path)) Then
            For Each item In System.IO.Directory.GetDirectories(Target_Path)
                info = System.IO.Path.GetDirectoryName(item).ToUpper.Split("_")
                If (info(0) = "RESERVED") Then mReserveList.Add(System.IO.Path.GetFileName(item))
            Next
        End If

    End Sub

    ' dgp rev 3/6/09 Scan for any run matches
    Private Sub ScanMatches(ByVal path As String, ByVal name As String)

        Dim curarr = name.Split("_")
        Dim CurRun = curarr(curarr.Length - 1).ToUpper
        mMatchList = New ArrayList
        Dim info
        Dim item
        For Each item In System.IO.Directory.GetDirectories(path)
            info = System.IO.Path.GetFileName(item).ToUpper.Split("_")
            If (info(0) = "RESERVED" And info(1) = CurRun) Then mMatchList.Add(item)
            If (info(info.length - 1) = CurRun) Then mMatchList.Add(item)
        Next

    End Sub

    ' dgp rev 3/6/09 Remove any empty matches for given run
    Private Sub RemoveMatches()

        Dim name
        If (mMatchList.Count > 0) Then
            For Each name In mMatchList
                If (System.IO.Directory.exists(name)) Then
                    If (System.IO.Directory.GetFiles(name).Length = 0) Then System.IO.Directory.Delete(name)
                End If
            Next
        End If

    End Sub

    ' dgp rev 10/20/08 Return the full run name
    Private mFull_Name As String
    Public ReadOnly Property Full_Name() As String
        Get
            ' dgp rev 10/13/08 make sure assigned run is formated correctly - Rxxxxx
            Return AssignedData + "_" + AssignedMap + "_" + AssignedRun
        End Get
    End Property

    ' dgp rev 3/5/09 Build the target path for current data
    Private Function BuildTarget() As Boolean

        BuildTarget = False
        Target_Path = String.Format("\\{0}{1}", DataServer, Upload_Root)
        Target_Path = system.io.path.combine(Target_Path, AssignedUser)

        ScanMatches(Target_Path, Full_Name)

        Target_Path = system.io.path.combine(Target_Path, Full_Name)
        If (System.IO.Directory.exists(Target_Path)) Then
            If (MsgBox("Run already exists.  Overwrite?", MsgBoxStyle.YesNo) = MsgBoxResult.No) Then Exit Function
        End If
        BuildTarget = Utility.Create_Tree(Target_Path)

    End Function

    ' dgp rev 3/6/09 Clean Up Run
    Public Sub CleanUpRun()

        RemoveMatches()

    End Sub

    ' dgp rev 3/10/09 
    Private mTransferCount As Integer
    Public ReadOnly Property TransferCount() As Integer
        Get
            Return mTransferCount
        End Get
    End Property

    ' dgp rev 3/10/09 
    Private mSuccessfulCount As Integer
    Public ReadOnly Property SuccessfulCount() As Integer
        Get
            Return mSuccessfulCount
        End Get
    End Property

    ' dgp rev 3/10/09 
    Private mSourceExists As Boolean = False
    Private mSourcePath As String
    Public Property SourcePath() As String
        Get
            Return mSourcePath
        End Get
        Set(ByVal value As String)
            If System.IO.Directory.exists(value) Then
                mSourcePath = value
                mSourceExists = True
            End If
        End Set
    End Property
    ' dgp rev 7/11/07 upload data to server
    ' may need to use authentication
    ' dgp rev 7/27/07 Upload an FCS Run to Server
    ' dgp rev 3/4/09 Upload looks for the RESERVED run to replace
    ' dgp rev 3/5/09 Look for empty reserved folder and replace with data folder
    ' dgp rev 3/5/09 separate the building of the source and target from the transfer
    Public Sub UpLoad()

        ' dgp rev 3/10/09 make sure the source exists
        If (Not mSourceExists) Then Exit Sub
        mSuccessfulCount = 0

        objImp.ImpersonateStart()
        If (BuildTarget()) Then
            Dim SrcFil
            Dim DestFil As String
            mTransferCount = System.IO.Directory.GetFiles(SourcePath).Length
            RaiseEvent evtRepeat("Transfer...")
            For Each SrcFil In System.IO.Directory.GetFiles(SourcePath)
                RaiseEvent evtRepeat(System.IO.Path.GetFileName(SrcFil))
                DestFil = System.IO.Path.Combine(Target_Path, System.IO.Path.GetFileName(SrcFil))
                System.IO.File.Copy(SrcFil, DestFil, True)
                If System.IO.File.Exists(DestFil) Then mSuccessfulCount = mSuccessfulCount + 1
            Next
        End If
        FCSUpload.objImp.ImpersonateStop()

        CleanUpRun()

    End Sub


    ' dgp rev 10/16/08 Must have data and run before upload is ready
    ' Upload Criteria
    ' - not already uploaded
    ' - successfully cached
    ' - user selected (mapped to VMS)
    ' - run assigned 
    Private mAssignedRun As String
    Public Property AssignedRun() As String
        Get
            Return mAssignedRun
        End Get
        Set(ByVal value As String)
            mAssignedRun = value
            mRunSetFlag = True
        End Set
    End Property

    ' dgp rev 10/16/08 
    Private mAssignedUser As String
    Public Property AssignedUser() As String
        Get
            Return mAssignedUser
        End Get
        Set(ByVal value As String)
            mAssignedUser = value
            mUserSetFlag = True
        End Set
    End Property

    ' dgp rev 10/16/08 
    Private mAssignedMap As String
    Public Property AssignedMap() As String
        Get
            Return mAssignedMap
        End Get
        Set(ByVal value As String)
            mAssignedMap = value
            mMapSetFlag = True
        End Set
    End Property

    ' dgp rev 10/16/08 
    Private mAssignedData As String
    Public Property AssignedData() As String
        Get
            Return mAssignedData
        End Get
        Set(ByVal value As String)
            mAssignedData = value
            mDataSetFlag = True
        End Set
    End Property

    ' dgp rev 10/16/08 
    Private mUserSetFlag As Boolean = False
    Private mRunSetFlag As Boolean = False
    Private mMapSetFlag As Boolean = False
    Private mDataSetFlag As Boolean = False

    ' dgp rev 10/16/08 
    Public Function CreateMap(ByVal user As String) As Boolean

        objImp.ImpersonateStart()
        PCtoVMS.PutSetting(Me.AssignedUser, user)
        CreateMap = PCtoVMS.Exists(Me.AssignedUser)
        objImp.ImpersonateStop()

    End Function

    ' dgp rev 10/16/08 
    Public ReadOnly Property IsReady() As Boolean
        Get
            Return mUserSetFlag And mRunSetFlag And mMapSetFlag And mDataSetFlag
        End Get
    End Property

    ' dgp rev 10/16/08 
    Public ReadOnly Property RunSetFlag() As Boolean
        Get
            Return mRunSetFlag
        End Get
    End Property

    ' dgp rev 10/16/08 
    Public ReadOnly Property UserSetFlag() As Boolean
        Get
            Return mUserSetFlag
        End Get
    End Property

    ' dgp rev 10/16/08 
    Public ReadOnly Property MapSetFlag() As Boolean
        Get
            Return mMapSetFlag
        End Get
    End Property

    ' dgp rev 10/16/08 
    Public ReadOnly Property DataSetFlag() As Boolean
        Get
            Return mDataSetFlag
        End Get
    End Property

    ' dgp rev 8/1/08 User Run has been selected
    Public Sub Run_User_Select(ByVal PCUser As String)

        If (PCtoVMS.Exists(PCUser)) Then
            AssignedMap = PCtoVMS.GetSetting(PCUser)
            '            Read_Count()
            RaiseEvent evtRepeat(PCUser + " maps to " + AssignedMap)
        Else
            RaiseEvent evtRepeat(PCUser + " must map to VMS account")
        End If

    End Sub

    ' dgp rev 11/1/07 Read File
    ' dgp rev 11/1/07 Last Run is stored in VMS named text file
    ' cross referenced with PC name via XML file 
    Public Sub Read_Count()

        FCSUpload.objImp.ImpersonateStart()
        Last_Run_Num = 0
        Dim VMSNum = 0
        Dim Run_File As String = system.io.path.combine(FTP_Path, AssignedMap + ".vms")
        If (System.IO.file.exists(Run_File)) Then
            Dim sr As New IO.StreamReader(Run_File)
            Dim raw As String = sr.ReadToEnd
            sr.Close()
            Try
                Last_Run_Num = CInt(raw)
            Catch ex As Exception
            End Try
        End If
        Dim XMLNum = 0

        objImp.ImpersonateStop()

    End Sub


    ' dgp rev 1/18/07 
    Private Function Check_Server_Access() As Boolean

        If (VMSFlag) Then
            If (Dyn.Exists("VMSServer")) Then
                DataServer = Dyn.GetSetting("VMSServer")
            Else
                DataServer = "spiffy.nci.nih.gov"
            End If
        Else
            If (Dyn.Exists("WinServer")) Then
                DataServer = Dyn.GetSetting("WinServer")
            Else
                DataServer = "NT-EIB-10-6B16"
            End If
        End If

        Try
            ServerOn = My.Computer.Network.Ping(DataServer, 1000)
            ' dgp rev 3/5/09 switch over to a global mapping
            PCtoVMS = New Dynamic(RemoteSettingsPath, "PCtoVMS")
        Catch ex As Exception
            ServerOn = False
        End Try

        Return ServerOn

    End Function

    ' dgp rev 7/9/07 Initial Program 
    ' check server access and find the data depot
    Public Function Init() As Boolean

        Return Check_Server_Access()

    End Function

End Class
