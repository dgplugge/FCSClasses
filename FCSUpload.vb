' Name:     FCS Server Download Routines
' Author:   Donald G Plugge
' Date:     1/26/07
' Purpose:  Class to handle operator downloads to the server
Imports HelperClasses
Imports FCS_Classes
Imports System.Management
Imports System.Threading
Imports System.Xml.Linq
Imports System.Xml

Public Class FCSUpload

    Private Property BackupFlag As Boolean = True

    Private Property EmailFlag As Boolean

    ' use delegates for Repeat
    Delegate Sub UploadFileEvent(ByVal Xname As String)
    Public Event UploadEventHandler As UploadFileEvent

    ' use delegates for Repeat
    Delegate Sub MainTransferDoneEvent(ByVal State As Boolean)
    Public Event MainTransferDoneEventHandler As MainTransferDoneEvent

    Delegate Sub BackupTransferDoneEvent(ByVal State As Boolean)
    Public Event BackupTransferDoneEventHandler As BackupTransferDoneEvent

    Delegate Sub TransferDoneEvent()
    Public Event TransferDoneEventHandler As TransferDoneEvent

    Public Shared objImp As New RunAs_Impersonator
    Public objFCS_File As FCS_Classes.FCS_File
    Public objVMS As VMSAccess


    ' dgp rev 11/10/2011
    Private mMainServerPath As String
    Public ReadOnly Property Target_Path As String
        Get
            Return mMainServerPath
        End Get
    End Property

    ' dgp rev 11/10/2011
    Private mBackup_Path As String
    Public ReadOnly Property Backup_Path As String
        Get
            Return mBackup_Path
        End Get
    End Property

    Private Shared mServerUploadRoot As String = "\Upload\FCSRun"

    ' dgp rev 11/10/2011
    Public Shared ReadOnly Property Upload_Root As String
        Get
            Return mServerUploadRoot
        End Get
    End Property

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
            If (System.IO.Directory.Exists(value)) Then
                mStoreRoot = value
                mValidStore = True
            End If
        End Set
    End Property

    ' dgp rev 11/10/2011
    Private mStoredFlag As Boolean = False
    Public ReadOnly Property StoredFlag() As Boolean
        Get
            Return mStoredFlag
        End Get
    End Property

    ' dgp rev 3/11/09 Valid Run Store
    Public ReadOnly Property StoreRunPath() As String
        Get
            If (ValidStore) Then Return System.IO.Path.Combine(StoreRoot, "FCSRun")
            Return ""
        End Get
    End Property

    ' dgp rev 3/11/09 Valid Run Store
    Public ReadOnly Property StoreExperPath() As String
        Get
            If (ValidStore) Then Return System.IO.Path.Combine(StoreRoot, "FCSExper")
            Return ""
        End Get
    End Property

    ' dgp rev 11/10/2011
    Private mOrigPath As String
    Public ReadOnly Property OrigPath() As String
        Get
            Return mOrigPath
        End Get
    End Property

    ' dgp rev 8/1/08
    Public Function Store_Files() As Boolean

        Store_Files = False
        If (System.IO.Directory.Exists(OrigPath)) Then
            If (System.IO.Directory.Exists(StoreRoot)) Then
                Dim path = System.IO.Path.Combine(StoreRoot, Format(Now(), "yyyyMMdd"))
                path = System.IO.Path.Combine(path, Me.AssignedUser)
                If Utility.Create_Tree(path) Then
                    path = System.IO.Path.Combine(path, Full_Name)
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
        If (System.IO.Directory.Exists(OrigPath)) Then
            If (System.IO.Directory.Exists(StoreExperPath)) Then
                Dim path = System.IO.Path.Combine(StoreExperPath, Format(Now(), "yyyyMMdd"))
                path = System.IO.Path.Combine(path, Me.AssignedUser)
                If Utility.Create_Tree(path) Then
                    path = System.IO.Path.Combine(path, Full_Name)
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
            mMainServerPath = AriaRunRoot
        Else
            mMainServerPath = String.Format("\\{0}{1}", DataServer, Upload_Root)
        End If
        mMainServerPath = System.IO.Path.Combine(mMainServerPath, AssignedUser)

        mReserveList = New ArrayList
        Dim info
        Dim item
        If (System.IO.Directory.Exists(mMainServerPath)) Then
            For Each item In System.IO.Directory.GetDirectories(mMainServerPath)
                info = System.IO.Path.GetFileName(item).ToUpper.Split("_")
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
                If (System.IO.Directory.Exists(name)) Then
                    If (System.IO.Directory.GetFiles(name).Length = 0) Then Utility.DeleteTree(name)
                End If
            Next
        End If

    End Sub

    ' dgp rev 10/20/08 Return the full run name
    Private mFull_Name As String
    Public ReadOnly Property Full_Name() As String
        Get
            If (mMapSetFlag) Then
                ' dgp rev 10/13/08 make sure assigned run is formated correctly - Rxxxxx
                Return AssignedData + "_" + AssignedMap + "_" + AssignedRun
            Else
                ' dgp rev 10/13/08 make sure assigned run is formated correctly - Rxxxxx
                Return AssignedData + "_" + AssignedUser + "_" + AssignedRun
            End If
        End Get
    End Property

    ' dgp rev 3/5/09 Build the target path for current data
    Private mBackupReady As Boolean = False
    Public ReadOnly Property BackupReady As Boolean
        Get
            Return mBackupReady
        End Get
    End Property

    ' dgp rev 3/5/09 Build the target path for current data
    Private mMainPathReady As Boolean = False
    Public ReadOnly Property MainPathReady As Boolean
        Get
            Return mMainPathReady
        End Get
    End Property

    ' dgp rev 8/30/2011 Scan Backup Path
    Public Function DeleteBackupPath() As Boolean

        mBackup_Path = String.Format("\\ncifs-p022.nci.nih.gov\Group\EIB\Branch99\FCSAria\")
        mBackup_Path = System.IO.Path.Combine(mBackup_Path, AssignedUser)

        mBackup_Path = System.IO.Path.Combine(mBackup_Path, Full_Name)
        mBackupReady = System.IO.Directory.Exists(mBackup_Path)
        Backup_Alias = mBackup_Path.Replace("\\ncifs-p022.nci.nih.gov\Group\EIB\Branch99", "I:")
        DeleteBackupPath = System.IO.Directory.Exists(mBackup_Path)
        mBackupReady = Utility.Create_Tree(mBackup_Path)
        If System.IO.Directory.Exists(mBackup_Path) Then
            Try
                System.IO.Directory.Delete(mBackup_Path, True)
            Catch ex As Exception

            End Try
        End If
        Return Not System.IO.Directory.Exists(mBackup_Path)

    End Function



    ' dgp rev 8/30/2011 Delete the main path
    Public Function DeleteMainPath() As Boolean

        DeleteMainPath = False
        mMainServerPath = String.Format("\\{0}{1}", DataServer, Upload_Root)
        mMainServerPath = System.IO.Path.Combine(mMainServerPath, AssignedUser)

        ScanMatches(mMainServerPath, Full_Name)

        mMainServerPath = System.IO.Path.Combine(mMainServerPath, Full_Name)
        If System.IO.Directory.Exists(mMainServerPath) Then
            Try
                System.IO.Directory.Delete(mMainServerPath, True)
            Catch ex As Exception

            End Try
        End If
        Return Not System.IO.Directory.Exists(mMainServerPath)

    End Function

    ' dgp rev 3/5/09 Build the target path for current data
    Public Function ScanMainPath() As Boolean

        ScanMainPath = False
        mMainServerPath = String.Format("\\{0}{1}", DataServer, Upload_Root)
        mMainServerPath = System.IO.Path.Combine(mMainServerPath, AssignedUser)

        ScanMatches(mMainServerPath, Full_Name)

        mMainServerPath = System.IO.Path.Combine(mMainServerPath, Full_Name)
        ScanMainPath = System.IO.Directory.Exists(mMainServerPath)
        mMainPathReady = Utility.Create_Tree(mMainServerPath)

    End Function

    ' dgp rev 3/6/09 Clean Up Run
    Public Sub CleanUpRun()

        RemoveMatches()

    End Sub

    ' dgp rev 3/10/09 
    Private mTransferCount As Integer

    ' dgp rev 3/10/09 
    Private mSuccessfulCount As Integer
    Public ReadOnly Property SuccessfulCount() As Integer
        Get
            Return mSuccessfulCount
        End Get
    End Property

    Private mTransferPath As String
    Private mBackupPath As String

    ' dgp rev 9/15/2011 Notify User
    Private Sub NotifyComplete()

        MsgBox("Transfer Complete", MsgBoxStyle.Information)
        If Send_Location() Then
            MsgBox(String.Format("Email Sent to {0}", CacheUser), MsgBoxStyle.Information)
        End If

    End Sub

    Public objMess As System.Net.Mail.MailMessage
    Public objMail As System.Net.Mail.SmtpClient

    Public EmailText As String
    Public EmailRecipient As String

    Private mEmailTo = "aria@mail.nih.gov"
    Private mMailServer = "mailfwd.nih.gov"
    Private mMailPort = 25

    ' dgp rev 10/12/07 
    ' dgp rev 10/12/07 
    Public Sub Prep_Message()

        If (System.Diagnostics.Debugger.IsAttached) Then
            EmailRecipient = "plugged@mail.nih.gov"
        Else
            If (CacheUserExists) Then
                EmailRecipient = CacheUser + "@mail.nih.gov"
            Else
                EmailRecipient = "plugged@mail.nih.gov"
            End If
        End If

        objMess = New System.Net.Mail.MailMessage(mEmailTo, EmailRecipient, "FCS Data Ready", Local_Message)
        objMail = New System.Net.Mail.SmtpClient(mMailServer, mMailPort)

        objMail.UseDefaultCredentials = True

    End Sub

    ' dgp rev 8/2/07 Send the log files in a message
    Public Function Send_Location() As Boolean

        Send_Location = True
        Prep_Message()

        Try
            objMail.Send(objMess)
            'Log_Info("Message Sent to User " + PCUser)
        Catch ex As System.Net.Mail.SmtpException
            Send_Location = False
            'Log_Info("SMTP Error: " + ex.Message)
        Catch ex As Exception
            Send_Location = False
            'Log_Info(ex.Message)
        End Try

    End Function

    ' dgp rev 9/15/2011 Email User
    Private Sub EmailUser()

        Dim objReport As New EmailReporting

        '        Dim smtp As New SimpleMail.SMTPClient
        '       Dim mail As New SimpleMail.SMTPMailMessage
        '        With mail
        '        objReport = My.Settings.ExceptionEmail.ToString
        '        If _blnHaveException Then
        '
        '       objReport = "Handled Exception notification - " & _strExceptionType
        '      Else
        '     .Subject = "HandledExceptionManager notification"
        '    End If
        objReport.EmailText = Local_Message()
        ' End With
        '-- try to send email, but we don't care if it succeeds (for now)
        Try
            objReport.SendReport()

        Catch e As Exception
            Debug.WriteLine("** SMTP email failed to send!")
            Debug.WriteLine("** " & e.Message)
        End Try

    End Sub

    ' dgp rev 8/31/2011 Transfer with backup if required
    Private Sub TransferThread()

        MainTransferThread()
        If (BackupFlag) Then
            BackupTransferThread()
        End If
        NotifyComplete()

        RaiseEvent TransferDoneEventHandler()

    End Sub

    ' dgp rev 8/29/2011 Upload the Run, with a daisy chain to Backup
    Public Sub Upload_Run()

        '            Add code to run as UserName here 'everything between ImpersonateStart and ImpersonateStop will be run as the impersonated user
        'FlowServer.DefineLog(CacheRun.NCIUser, CacheRun.RunName)

        Dim objThread As New Thread(New ThreadStart(AddressOf TransferThread))
        objThread.Name = "FCS Main Transfer"
        objThread.Start()

        FlowServer.CloseLog()

    End Sub

    ' dgp rev 7/11/07 upload data to server
    ' may need to use authentication
    ' dgp rev 7/27/07 Upload an FCS Run to Server
    ' dgp rev 3/4/09 Upload looks for the RESERVED run to replace
    ' dgp rev 3/5/09 Look for empty reserved folder and replace with data folder
    ' dgp rev 3/5/09 separate the building of the source and target from the transfer

    ' dgp rev 10/12/07 Backup FCS Data
    Public Sub BackupTransferThread()

        Good_Backup = 0
        Dim SrcFil
        Dim DestFil As String

        If mBackupReady Then
            If (CacheDataExists()) Then
                If Utility.Create_Tree(mBackup_Path) Then
                    Total_Attempts = System.IO.Directory.GetFiles(CacheUserPath).Length
                    Try
                        RaiseEvent UploadEventHandler("Backup...")
                        For Each SrcFil In System.IO.Directory.GetFiles(CacheUserPath)
                            RaiseEvent UploadEventHandler(System.IO.Path.GetFileName(SrcFil))
                            DestFil = System.IO.Path.Combine(Backup_Path, System.IO.Path.GetFileName(SrcFil))
                            System.IO.File.Copy(SrcFil, DestFil, True)
                            If System.IO.File.Exists(DestFil) Then Good_Backup = Good_Backup + 1
                        Next
                    Catch ex As Exception
                        mErrorMessage = ex.Message
                        RaiseEvent BackupTransferDoneEventHandler(False)
                    End Try
                End If
            End If
            RaiseEvent BackupTransferDoneEventHandler(Good_Backup = Total_Attempts)
        Else
            RaiseEvent BackupTransferDoneEventHandler(False)
        End If

    End Sub

    ' dgp rev 11/10/2011
    Public ReadOnly Property TransferCount As Int16
        Get
            If System.IO.Directory.Exists(CacheUserPath) Then
                Return System.IO.Directory.GetFiles(CacheUserPath).Length
            End If
            Return 0
        End Get
    End Property

    ' dgp rev 8/30/2011 Main Transfer 
    Public Sub MainTransferThread()

        ' dgp rev 3/10/09 make sure the source exists
        If MainPathReady Then
            If (CacheDataExists) Then
                mSuccessfulCount = 0

                objImp.ImpersonateStart()
                Dim SrcFil
                Dim DestFil As String
                RaiseEvent UploadEventHandler("Transfer...")
                Try
                    For Each SrcFil In System.IO.Directory.GetFiles(CacheUserPath)
                        RaiseEvent UploadEventHandler(System.IO.Path.GetFileName(SrcFil))
                        DestFil = System.IO.Path.Combine(Target_Path, System.IO.Path.GetFileName(SrcFil))
                        System.IO.File.Copy(SrcFil, DestFil, True)
                        If System.IO.File.Exists(DestFil) Then mSuccessfulCount = mSuccessfulCount + 1
                    Next
                Catch ex As Exception
                    mErrorMessage = ex.Message
                    RaiseEvent MainTransferDoneEventHandler(False)
                    Exit Sub
                End Try
                FCSUpload.objImp.ImpersonateStop()
                RaiseEvent MainTransferDoneEventHandler(mSuccessfulCount = TransferCount)
            Else
                mErrorMessage = "Server not ready"
                RaiseEvent MainTransferDoneEventHandler(False)
            End If
        Else
            mErrorMessage = "Server not ready"
            RaiseEvent MainTransferDoneEventHandler(False)
        End If
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
            ' dgp rev 5/27/09 remove the mapping requirement
            Return mUserSetFlag And mRunSetFlag And mDataSetFlag
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

        If PCtoVMS IsNot Nothing Then
            If (PCtoVMS.Exists(PCUser)) Then
                AssignedMap = PCtoVMS.GetSetting(PCUser)
                '            Read_Count()
                RaiseEvent UploadEventHandler(PCUser + " maps to " + AssignedMap)
            Else
                RaiseEvent UploadEventHandler(PCUser + " must map to VMS account")
            End If
        End If

    End Sub

    ' dgp rev 11/1/07 Read File
    ' dgp rev 11/1/07 Last Run is stored in VMS named text file
    ' cross referenced with PC name via XML file 
    Public Sub Read_Count()

        FCSUpload.objImp.ImpersonateStart()
        Last_Run_Num = 0
        Dim VMSNum = 0
        Dim Run_File As String = System.IO.Path.Combine(FTP_Path, AssignedMap + ".vms")
        If (System.IO.File.Exists(Run_File)) Then
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

    ' dgp rev 10/06/09 
    Public Enum WorkingState
        StateError = -1
        [Nothing] = 0
        Data = 1
        User = 2
        Both = 3
    End Enum

    ' dgp rev 10/06/09 
    Private mCacheRoot
    Public ReadOnly Property CacheRoot() As String
        Get
            If (mCacheRoot Is Nothing) Then mCacheRoot = System.IO.Path.Combine(FlowStructure.FlowRoot, "Cache")
            Return mCacheRoot
        End Get

    End Property

    ' dgp rev 10/06/09 
    Public ReadOnly Property CacheRootExists() As Boolean
        Get
            If (CacheRoot Is Nothing) Then Return False
            Return System.IO.Directory.Exists(CacheRoot)
        End Get
    End Property

    ' dgp rev 10/06/09 
    Public ReadOnly Property CacheUserExists()
        Get
            If (Not CacheRootExists) Then Return False
            Return (System.IO.Directory.GetDirectories(CacheRoot).Length > 0)
        End Get
    End Property

    ' dgp rev 10/06/09 
    Public ReadOnly Property CacheUserPath()
        Get
            If (Not CacheUserExists) Then Return Nothing
            Dim User = System.IO.Directory.GetDirectories(CacheRoot)(0)
            Return System.IO.Path.Combine(CacheRoot, User)
        End Get
    End Property

    ' dgp rev 10/06/09 
    Public ReadOnly Property CacheDataExists() As Boolean
        Get
            If (Not CacheRootExists) Then Return False
            Dim DataPath
            If (System.IO.Directory.GetDirectories(CacheRoot).Length = 0) Then
                DataPath = CacheRoot
            Else
                DataPath = CacheUserPath
            End If
            Return (System.IO.Directory.GetFiles(DataPath).Length > 0)

        End Get
    End Property

    ' dgp rev 10/06/09 
    Public ReadOnly Property CacheDataPath()
        Get
            If (Not CacheRootExists) Then Return ""
            If (System.IO.Directory.GetDirectories(CacheRoot).Length = 0) Then
                Return CacheRoot
            Else
                Return CacheUserPath
            End If

        End Get
    End Property

    ' dgp rev 10/06/09 
    Public Function CheckState() As WorkingState

        If (Not CacheRootExists) Then Return WorkingState.Nothing
        If (CacheUserExists) Then
            If (CacheDataExists) Then Return WorkingState.Both
            Return WorkingState.User
        Else
            If (CacheDataExists) Then Return WorkingState.Data
            Return WorkingState.Nothing
        End If

    End Function

    ' dgp rev 11/3/08 Color of Current Stage
    Public ReadOnly Property StageColor() As System.Drawing.Color
        Get
            Select Case CheckState()
                Case WorkingState.Nothing
                    Return Drawing.Color.White
                Case WorkingState.Both
                    Return Drawing.Color.GreenYellow
                Case WorkingState.Data
                    Return Drawing.Color.LightGoldenrodYellow
                Case WorkingState.User
                    Return Drawing.Color.LightGoldenrodYellow
            End Select
        End Get
    End Property

    ' dgp rev 11/10/2011
    Private Sub AppendLog(ByVal txt As String)

        RaiseEvent UploadEventHandler(txt)
        mLogInfo = mLogInfo + vbCrLf + txt

    End Sub

    ' dgp rev 11/10/2011
    Private Sub Log_Exper()

        AppendLog("Valid Experiment")
        AppendLog(mUniqueName)
        AppendLog(mLocalData.Experiment)
        AppendLog(mLocalData.Run)

    End Sub

    Private Shared mUniqueName As String
    ' server
    Public Exper_Root As String = "\Upload\FCSExper"
    Private Status As String

    ' dgp rev 6/21/07 Evaluate the Selected Path
    Public Sub Evaluate_Path(ByVal path As String)

        ' create an experiment, then assign the depot location
        mLocalData = Nothing
        mLocalData = New FCSRun(path)
        '        objUploadRun.Depot_Root = Depot_Root

        ' dgp rev 8/6/08 Only one instance, experiment or run
        'objUploadRun = Nothing
        'objUploadRun = New FCS_Classes.FCSRun(path)
        '            Users = objUploadRun.Users

        If (mLocalData.Valid_Run) Then
            '                mForceUser = objUploadRun.Users.Count > 0
            CacheData()
            mCacheRun = New FCSRun(CacheDataPath)
            mCacheRun.CreateRealList()
            If (CacheUserExists) Then NCIUser = CacheUser
            AppendLog(Status)
            ' dgp rev 9/16/2011 xyzzy is upload_root neccessary?
            ' dgp rev 9/16/2011 xyzzy run root is fixed.
            ' Option 2 Add the event with the handler address
            AddHandler mCacheRun.evtRepeat, AddressOf AppendLog
            mCurType = FCSType.Run
            mUniqueName = mCacheRun.Unique_Prefix
            AssignedData = mUniqueName
            AppendLog("Valid Run")
            AppendLog(CStr(mCacheRun.FCS_cnt) + " FCS files")
            AppendLog(mUniqueName)
        Else
            mCurType = FCSType.Invalid
            AppendLog("Invalid Path")
        End If

        '        If (mForceUser) Then If (lstNCIUsers.Items.Contains(Users.Item(0))) Then lstNCIUsers.SelectedIndex = lstNCIUsers.Items.IndexOf(Users.Item(0))

    End Sub

    ' dgp rev 11/10/2011
    Private Shared mCacheRun As FCS_Classes.FCSRun
    Public Shared ReadOnly Property CacheRun As FCSRun
        Get
            Return mCacheRun
        End Get
    End Property


    ' dgp rev 11/10/2011
    Private Shared mLocalData As FCS_Classes.FCSRun
    Public Shared ReadOnly Property LocalData As FCSRun
        Get
            Return mLocalData
        End Get
    End Property

    ' dgp rev 11/10/2011
    Private Shared mCacheFlag As Boolean = True
    Public Shared ReadOnly Property CacheFlag As Boolean
        Get
            Return mCacheFlag
        End Get
    End Property


    ' dgp rev 10/06/09 
    Private Function CacheOriginalData() As Boolean

        CacheOriginalData = False
        If (Utility.Create_Tree(CacheRoot)) Then
            Dim item
            Try
                For Each item In System.IO.Directory.GetFiles(mLocalData.Orig_Path)
                    System.IO.File.Copy(item, System.IO.Path.Combine(CacheDataPath, System.IO.Path.GetFileName(item)))
                Next
                CacheOriginalData = True
            Catch ex As Exception
            End Try
        End If
        Return CacheOriginalData

    End Function

    ' dgp rev 10/06/09 
    Public Function RescanCache() As Boolean

        RescanCache = False
        If System.IO.Directory.Exists(CacheUserPath) Then
            mCacheRun = Nothing
            mCacheRun = New FCSRun(CacheUserPath)
            RescanCache = True
        End If

    End Function

    ' dgp rev 10/06/09 
    Private Function MoveCache() As Boolean

        Dim item
        Try
            For Each item In System.IO.Directory.GetFiles(CacheRoot)
                System.IO.File.Move(item, System.IO.Path.Combine(CacheUserPath, System.IO.Path.GetFileName(item)))
            Next
            mCacheRun = Nothing
            mCacheRun = New FCSRun(CacheUserPath)
            Return True
        Catch ex As Exception
            Return False
        End Try


    End Function

    ' dgp rev 10/06/09 
    Private Function RenameUser(ByVal user) As Boolean

        If user = CacheUser Then Return True
        Try
            System.IO.Directory.Move(CacheUserPath, System.IO.Path.Combine(CacheRoot, user))
            mCacheRun = Nothing
            mCacheRun = New FCSRun(CacheUserPath)
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 10/06/09 
    Public ReadOnly Property CacheUser()
        Get
            If (Not CacheUserExists) Then Return Nothing
            Return System.IO.Path.GetFileNameWithoutExtension(System.IO.Directory.GetDirectories(CacheRoot)(0))
        End Get
    End Property

    ' dgp rev 10/06/09 
    Private Function CreateUser(ByVal user) As Boolean

        If (CacheUserExists) Then
            If (user = CacheUser) Then Return True
            RenameUser(user)
        Else
            Return Utility.Create_Tree(System.IO.Path.Combine(CacheRoot, user))
        End If

    End Function

    ' dgp rev 10/06/09 
    Private Function AssignUser(ByVal user) As Boolean

        If (CreateUser(user)) Then
            MoveCache()
        End If

    End Function

    ' dgp rev 11/10/2011
    Private Shared mLogInfo As String
    Public Shared ReadOnly Property LogInfo
        Get
            Return mLogInfo
        End Get
    End Property

    Public Good_Backup As Int16
    Public Good_Upload As Int16
    Public Backup_Alias As String
    Public Total_Attempts As Int16

    ' dgp rev 8/30/2011 Scan Backup Path
    Public Function ScanBackupPath() As Boolean

        mBackup_Path = String.Format("\\ncifs-p022.nci.nih.gov\Group\EIB\Branch99\FCSAria\")
        mBackup_Path = System.IO.Path.Combine(mBackup_Path, AssignedUser)

        mBackup_Path = System.IO.Path.Combine(mBackup_Path, Full_Name)
        mBackupReady = System.IO.Directory.Exists(mBackup_Path)
        Backup_Alias = mBackup_Path.Replace("\\ncifs-p022.nci.nih.gov\Group\EIB\Branch99", "I:")
        mBackupReady = System.IO.Directory.Exists(mBackup_Path)
        Return (mBackupReady)

    End Function

    ' dgp rev 11/10/2011
    Private mErrorMessage As String = ""
    Public ReadOnly Property ErrorMessage As String
        Get
            Return mErrorMessage
        End Get
    End Property

    ' dgp rev 3/11/09 Remove CurRun, perhaps a valid backup must be confirmed
    Public Function Remove_Original() As Boolean

        ' dgp rev 7/29/08 in order to create the CurCache we need a valid experiment
        ' and a valid root
        Remove_Original = False
        If (CacheValidate(NCIUser)) Then
            Try
                Utility.DeleteTree(OrigPath)
                Remove_Original = True
            Catch ex As Exception
            End Try
        End If

    End Function

    ' dgp rev 11/10/2011
    Public Enum Stage
        StageError = -1
        [Nothing] = 0
        DataOnly = 1
        UserOnly = 2
        Ready = 3
        Uploaded = 4
        Stored = 5
    End Enum

    ' dgp rev 11/10/2011
    Public Enum StageError
        Invalid = -1
        Mismatch = -2
        CacheError = -3
        [Nothing] = 0
    End Enum

    ' dgp rev 10/2/09 
    Public Function CurStage() As Stage

        If (Not CacheRootExists) Then Return Stage.Nothing
        If (CacheUserExists) Then
            If (CacheDataExists) Then Return Stage.Ready
            Return Stage.UserOnly
        Else
            If (CacheDataExists) Then Return Stage.DataOnly
            Return Stage.Nothing
        End If

    End Function

    ' dgp rev 6/1/09 Folder contains only FCS files
    Private Function AllFCSSub(ByVal path) As Boolean

        AllFCSSub = False
        Dim item
        Dim objFCS As FCS_File
        If System.IO.Directory.Exists(path) Then
            For Each item In System.IO.Directory.GetFiles(path)
                objFCS = New FCS_File(item)
                If (Not objFCS.Valid) Then Exit Function
            Next
            AllFCSSub = True
        End If

    End Function

    Private mStatus As String
    ' dgp rev 6/26/07 Change the Experiment Location
    ' move the selected experiment into a local cache in the flowroot depot
    ' then add a checksum for validatio
    Private Function CacheValidate(ByVal user As String) As Boolean

        CacheValidate = False
        mStatus = "Directory Failure"

        If (CacheDataExists) Then
            If (AllFCSSub(CacheDataPath)) Then Return True
        End If

    End Function

    ' dgp rev 10/06/09 
    Public Function ClearCache()

        Try
            Utility.DeleteTree(CacheRoot)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    ' dgp rev 10/06/09 
    Private Function ReplaceUserData()

        Try
            Dim item
            For Each item In System.IO.Directory.GetFiles(System.IO.Path.Combine(CacheRoot, CacheUser))
                System.IO.File.Delete(item)
            Next
            For Each item In System.IO.Directory.GetFiles(OrigPath)
                System.IO.File.Copy(item, System.IO.Path.Combine(CacheUserPath, System.IO.Path.GetFileName(item)))
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    ' dgp rev 10/06/09 
    Public Sub CacheData()

        Select Case CheckState()
            Case WorkingState.Nothing
                CacheOriginalData()
            Case WorkingState.User
                CacheOriginalData()
                AssignUser(CacheUser)
            Case WorkingState.Data
                ClearCache()
                CacheOriginalData()
            Case WorkingState.Both
                ReplaceUserData()
        End Select

    End Sub

    ' dgp rev 10/16/08
    Public ReadOnly Property Successful_Upload() As Boolean
        Get
            Return (TransferCount = SuccessfulCount)
        End Get
    End Property

    Public Delegate Sub InfoEventHandler(ByVal Message As String)
    Public Shared Event InfoEvent As InfoEventHandler

    Private mCheckSumDoc As XDocument
    Private mExperElement As XElement

    ' dgp rev 11/10/2011
    Private mFCSRunUploadLogName = Nothing
    Public ReadOnly Property FCSRunUploadLogName() As String
        Get
            If (mFCSRunUploadLogName Is Nothing) Then mFCSRunUploadLogName = String.Format("{0}.xml", Now.ToLongDateString)
            Return mFCSRunUploadLogName
        End Get
    End Property

    ' dgp rev 11/10/2011
    Private mFCSRunUploadLogPath = Nothing
    Public ReadOnly Property FCSRunUploadLogPath() As String
        Get
            If (mFCSRunUploadLogPath Is Nothing) Then mFCSRunUploadLogPath = System.IO.Path.Combine(FlowStructure.FlowRoot, "Upload\LogFiles")
            Return mFCSRunUploadLogPath
        End Get
    End Property

    ' dgp rev 11/10/2011
    Public ReadOnly Property FCSRunUploadLogFullSpec As String
        Get
            Return System.IO.Path.Combine(FCSRunUploadLogPath, FCSRunUploadLogName)
        End Get
    End Property

    ' dgp rev 10/19/2010
    Private Function CreateLogElement() As XElement

        Try
            CreateLogElement =
                New XElement("Experiment",
                                  New XElement("Ident", New XAttribute("modified", DateTime.Now)),
                                  New XElement("Unique", FCSUpload.mUniqueName),
                                  New XElement("ShareName", FlowServer.ShareExperRoot),
                                  New XElement("ShareInfo", FCSUpload.mServerUploadRoot),
                                  New XElement("Relative", Target_Path),
                                  New XElement("User", AssignedUser)
                                )

        Catch ex As Exception
            mMessage = ex.Message
            CreateLogElement =
                New XElement("Experiment",
                                  New XElement("Timestamp", New XAttribute("modified", DateTime.Now)),
                                  New XElement("Error", ex.Message),
                                  New XElement("Date", Now.Date.ToLongDateString))
        End Try


    End Function

    Private mMessage As String

    Private mMainTree As XElement
    Private mNewBranch As XElement
    ' dgp rev 5/31/2011
    Public Function AppendLog() As Boolean

        Dim LocalValid = False
        AppendLog = False
        If System.IO.File.Exists(FCSRunUploadLogFullSpec) Then
            Dim XMLDoc As XDocument = XDocument.Load(FCSRunUploadLogFullSpec)
            Dim found = From info In XMLDoc.Descendants("Experiment")
                        Where info.Elements("Unique").Value = mUniqueName 
            If found.Count = 0 Then
                mMainTree = XElement.Load(FCSRunUploadLogFullSpec)
                Dim firstExper As XElement = mMainTree.Element("Experiment")
                firstExper.AddBeforeSelf(mExperElement)
                LocalValid = True
            End If
        Else
            mMainTree = New XElement("Experiments", mExperElement)
            LocalValid = True
        End If

    End Function



    Private mLogDoc As XDocument
    ' dgp rev 9/7/2011 Update the log file
    Private Function CreateLogFile() As Boolean

        Dim LogElement As XElement = CreateLogElement()
        mLogDoc =
            New XDocument(
                New XDeclaration("1.0", "utf-8", "yes"), LogElement)
        If Utility.Create_Tree(FCSRunUploadLogPath) Then
            mLogDoc.Save(FCSRunUploadLogFullSpec)
        End If
        Try

        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    ' dgp rev 5/28/2011
    Public Function RecordUpload() As Boolean

        RaiseEvent InfoEvent(String.Format("{0}", "Create Log"))
        Return CreateLogFile()
        Return False

    End Function



    ' dpg rev 10/16/08 check the upload status and if successful, clear cache
    Public Function CheckNClear() As Boolean

        ClearCache()
        RecordUpload()

        Return (Successful_Upload)

    End Function

    ' dgp rev 10/06/09 
    Public Sub ProcessNewUser(ByVal user)

        Select Case CheckState()
            Case WorkingState.Nothing
                CreateUser(user)
            Case WorkingState.User
                RenameUser(user)
            Case WorkingState.Data
                AssignUser(user)
            Case WorkingState.Both
                RenameUser(user)
        End Select

    End Sub

    ' dgp rev 11/10/2011
    Public Enum FCSType
        Invalid = -1
        [Nothing] = 0
        Experiment = 1
        Run = 2
    End Enum

    ' dgp rev 11/10/2011
    Private mCurType As FCSType
    Public ReadOnly Property CurType As FCSType
        Get
            Return mCurType
        End Get
    End Property

    ' dgp rev 8/1/08
    Public Function Store_Away() As Boolean

        Select Case CurType
            Case FCSType.Experiment
                If (CurStage() = Stage.Uploaded) Then
                    Store_Away = Store_Exper()
                    mStoredFlag = Store_Away
                    If (Remove_Original()) Then

                    End If
                End If
            Case FCSType.Run
                If (CurStage() = Stage.Uploaded) Then
                    Store_Away = Store_Files()
                    mStoredFlag = Store_Away
                End If
        End Select

    End Function
    ' dgp rev 6/3/09
    ' dgp rev 10/6/09 
    Private mNCIUser As String
    Public Property NCIUser() As String
        Get
            Return mNCIUser
        End Get
        Set(ByVal value As String)
            If (value Is Nothing) Then
                ClearCache()
                Exit Property
            End If
            If (mNCIUser IsNot Nothing) Then
                If (mNCIUser.ToString = value) Then Exit Property
            End If
            mNCIUser = value
            ProcessNewUser(value)
        End Set
    End Property

    ' dgp rev 7/9/07 Initial Program 
    ' check server access and find the data depot
    Public Function Init() As Boolean

        Return Check_Server_Access()

    End Function

    Private Sub TransferStatus(ByVal Xname As String)
        Throw New NotImplementedException
    End Sub

    ' dgp rev 8/29/2011 Remote Message
    Public Function Local_Message() As String

        Dim EmailText As String = "Your data has been copied to a network drive."
        EmailText = EmailText + vbCrLf + CacheUser
        ' dgp rev 9/16/2011 xyzzy cache run is gone by this point
        EmailText = EmailText + vbCrLf + CacheRun.Unique_Prefix()
        EmailText = EmailText + vbCrLf
        EmailText = EmailText + vbCrLf + "Location file://" + Backup_Alias
        EmailText = EmailText + vbCrLf + "Location file:" + Backup_Path

        Return EmailText

    End Function

    ' dgp rev 8/29/2011 Remote Message
    Public Function Remote_Message() As String

        Dim EmailText = "Your data has been uploaded to the Flow Lab Server."
        EmailText = EmailText + vbCrLf + CacheUser
        EmailText = EmailText + vbCrLf + CacheRun.Unique_Prefix()
        EmailText = EmailText + vbCrLf
        EmailText = EmailText + vbCrLf + "Please use the FCS Aria utility to download the FCS files."
        EmailText = EmailText + vbCrLf
        EmailText = EmailText + vbCrLf + "This utility can be installed from http://flowroot.nci.nih.gov/distribution/FCSAriaDeploy.msi"
        EmailText = EmailText + vbCrLf
        EmailText = EmailText + vbCrLf + "For help and an overview click http://flowroot.nci.nih.gov/OpenWiki/ow.asp?AriaDownload"

        Return EmailText

    End Function


End Class
