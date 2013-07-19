' Name:     FlowServer
' Author:   Donald G Plugge
' Date:     11/14/08
' Purpose:  Class for interacting with the NT-EIB-10-6B16 Flow Lab Server

Imports System.IO
Imports HelperClasses
Imports System.Management
Imports System.Xml.Linq


Public Class FlowServer

    Public Shared XMLSettings As New Dynamic("FlowServer")

    Private Shared mServerOn As Boolean = False

    ' dgp rev 12/10/08 setup handles for event monitoring
    Public Delegate Sub FSSRemoteEventHandler(ByVal SomeString As String)
    Public Shared Event FSSRemoteEvent As FSSRemoteEventHandler

    Private Shared mServerName As String = "NT-EIB-10-6B16"
    Private Shared m_Remote_Registery As String = String.Format("\\{0}\upload\", mServerName)
    Private Shared m_ManagementPath As String = String.Format("\\{0}\root\cimv2:Win32_Group.Domain='{0}'", mServerName)

    ' dgp rev 7/1/09 
    Private Shared mAuthorized As Boolean = True
    Public Shared ReadOnly Property Authorized() As Boolean
        Get
            Return mAuthorized
        End Get
    End Property

    ' dgp rev 11/19/08 Impersonation memeber
    Private Shared mImpersonate As New HelperClasses.RunAs_Impersonator
    Public Shared Property Impersonate() As HelperClasses.RunAs_Impersonator
        Get
            Return mImpersonate
        End Get
        Set(ByVal value As HelperClasses.RunAs_Impersonator)
            mImpersonate = value
        End Set
    End Property

    Private Shared mFlowServer As String = mServerName
    Private Shared mCIMV2 As String
    Private Shared mScope As ManagementScope

    ' dgp rev 5/12/2011 Get Freespace on given drive
    Public Shared Function ServerPathExists(ByVal path As String) As Boolean

        Dim ms As System.Management.ManagementScope
        Dim oq As System.Management.ObjectQuery
        Dim mos As System.Management.ManagementObjectSearcher
        Dim obj As System.Management.ManagementObject

        ServerPathExists = False
        Try
            ' The "scope" includes the name of the PC and the WMI namespace
            ms = New System.Management.ManagementScope("\\" & FlowServer.ToString & _
             "\root\cimv2")

            ' use WQL to get just the one instance we want.  This should look familiar
            ' to those who are used to SQL
            path = path.Replace("\", "\\")
            oq = New System.Management.ObjectQuery(String.Format("select * from Win32_directory where Name = '{0}'", path))
            ' execute the query
            mos = New System.Management.ManagementObjectSearcher(ms, oq)
            For Each obj In mos.Get
                ServerPathExists = True
            Next
        Catch ex As Exception

        End Try

    End Function



    ' dgp rev 5/12/2011 Get Freespace on given drive
    Public Shared Function GetSpace(ByVal drv As String) As Integer

        Dim ms As System.Management.ManagementScope
        Dim oq As System.Management.ObjectQuery
        Dim mos As System.Management.ManagementObjectSearcher
        Dim obj As System.Management.ManagementObject
        Dim ans As Double

        ' The "scope" includes the name of the PC and the WMI namespace
        ms = New System.Management.ManagementScope("\\" & FlowServer.ToString & _
         "\root\cimv2")

        ' use WQL to get just the one instance we want.  This should look familiar
        ' to those who are used to SQL
        oq = New System.Management.ObjectQuery("select FreeSpace from Win32_LogicalDisk where DriveType=3 and DeviceId = '" + drv + "'")
        ' execute the query
        mos = New System.Management.ManagementObjectSearcher(ms, oq)
        For Each obj In mos.Get
            ans = Convert.ToDouble(obj("FreeSpace")) / (1024.0# * 1024.0#)
            Exit For
        Next

        Return ans

    End Function

    ' dgp rev 5/17/2011
    Private Shared mExperXMLShare As String = "ExperXMLShare"
    Public Shared ReadOnly Property ExperXMLShare
        Get
            Return mExperXMLShare
        End Get
    End Property

    Private mExperXMLRoot As String = "Experiments"
    Private mManWindowShare As String = "ManWindow"

    ' dgp rev 7/1/09 
    Public Shared ReadOnly Property FlowServer() As String
        Get
            Return mFlowServer
        End Get
    End Property
    ' dgp rev 11/14/08 Is the server up?
    Public Shared Function Server_Up() As Boolean

        If (mServerOn) Then Return True

        If (XMLSettings.Exists("FlowServer")) Then
            mFlowServer = XMLSettings.GetSetting("FlowServer")
        Else
            mFlowServer = mServerName
        End If

        Try
            mServerOn = My.Computer.Network.Ping(FlowServer, 1000)
        Catch ex As Exception
            mServerOn = False
        End Try

        Return mServerOn

    End Function

    ' dgp rev 12/5/08 Verified alternate username
    Private Shared mImperson As String = ""
    Public Shared Property Imperson() As String
        Get
            If (mImperson = "") Then mImperson = mUsername
            Return mImperson
        End Get
        Set(ByVal value As String)
            mImperson = value
        End Set
    End Property

    ' dgp rev 12/5/08 Actual logged on username
    Private Shared mUsername As String = ""
    Public Shared ReadOnly Property Username() As String
        Get
            If (mUsername = "") Then mUsername = Environment.GetEnvironmentVariable("username")
            Return mUsername
        End Get
    End Property
    ' dgp rev 11/17/08 Flow Root on the Flow Lab Server
    Private Shared mShareFlow As String = ""
    Public Shared ReadOnly Property ShareFlow() As String
        Get
            If (Not mShareFlow = "") Then Return mShareFlow
            If (XMLSettings.Exists("ShareFlow")) Then
                mShareFlow = XMLSettings.GetSetting("ShareFlow")
            Else
                mShareFlow = "root2"
            End If
            Return mShareFlow
        End Get
    End Property

    ' dgp rev 11/17/08 location on the server for user data
    Private Shared mUsersShare As String = ""
    Public Shared ReadOnly Property UsersShare() As String
        Get
            If (Not mUsersShare = "") Then Return mUsersShare
            If (XMLSettings.Exists("ShareFlow")) Then
                mUsersShare = XMLSettings.GetSetting("Users")
            Else
                mUsersShare = "Users"
            End If
            Return mUsersShare
        End Get
    End Property

    ' dgp rev 11/17/08 location on the server for user data
    Private Shared mUsersPath As String = ""
    Public Shared ReadOnly Property UsersPath() As String
        Get
            Return String.Format("\\{0}\{1}\", FlowServer, UsersShare)
        End Get
    End Property

    ' dgp rev 11/17/08 location on the server for user data
    Private Shared mUsersExists As Boolean = False
    Public Shared ReadOnly Property UsersExists() As Boolean
        Get
            mImpersonate.ImpersonateStart()
            Dim val = System.IO.Directory.exists(UsersPath)
            mImpersonate.ImpersonateStop()
            Return val
        End Get
    End Property

    ' dgp rev 11/17/08 location on the server for user data
    Private Shared mUserExists As Boolean = False
    Public Shared ReadOnly Property UserExists() As Boolean
        Get
            mImpersonate.ImpersonateStart()
            Dim val = System.IO.Directory.Exists(System.IO.Path.Combine(UsersPath, Username))
            mImpersonate.ImpersonateStop()
            Return val
        End Get
    End Property

    ' dgp rev 11/17/08 location on the server for user data
    Private Shared mUserPath As String
    Public Shared ReadOnly Property UserPath() As String
        Get
            ' dgp rev 6/15/09 Remote user path
            Return System.IO.Path.Combine(UsersPath, Username)
        End Get
    End Property

    ' dgp rev 11/17/08 location on the server for user data
    Private Shared mCreateUser As Boolean
    Public Shared ReadOnly Property CreateUser() As Boolean
        Get
            mImpersonate.ImpersonateStart()
            RaiseEvent FSSRemoteEvent("Check User Path")
            Dim val = Utility.Create_Tree(system.io.path.combine(UsersPath, Username))
            RaiseEvent FSSRemoteEvent("User Path Valid")
            mImpersonate.ImpersonateStop()
            Return val
        End Get
    End Property

    ' dgp rev 3/27/08 login information incase of priviledge issues
    Private Shared mwmiOptions As New ConnectionOptions

    ' dgp rev 7/1/09 
    Public Shared Property wmiOptions() As ConnectionOptions
        Get
            Return mwmiOptions
        End Get
        Set(ByVal value As ConnectionOptions)
            mwmiOptions = value
        End Set
    End Property

    ' dgp rev 11/14/08 Verify Authentication
    Public Shared Function Check_Admin() As Boolean

        ' dgp rev 3/27/08 Verify connection to server


        Check_Admin = True

        mCIMV2 = "\\" & FlowServer & "\root\cimv2:Win32_GroupUser"
        mScope = New ManagementScope(mCIMV2)

        '* connect to WMI namespace

        mImpersonate.ImpersonateStart()
        Try
            mScope.Connect()
            mAuthorized = True
        Catch ex As Exception
            Check_Admin = False
        End Try
        mImpersonate.ImpersonateStop()

    End Function

    ' dgp rev 7/1/09 
    Public Shared Property WMIPath() As String
        Get
            Return m_ManagementPath
        End Get
        Set(ByVal value As String)
            m_ManagementPath = value
        End Set
    End Property

    ' dgp rev 7/1/09 
    Public Shared Property Remote_Registery() As String
        Get
            Return m_Remote_Registery
        End Get
        Set(ByVal value As String)
            m_Remote_Registery = value
        End Set
    End Property

    ' dgp rev 10/13/2010
    Private Shared mTransferLog As String
    Private Shared mLoggingOn As Boolean = False
    Private Shared mLogWriter As StreamWriter

    Public Shared Sub WriteStatus(ByVal line As String)

        If mLoggingOn Then
            Try
                mLogWriter.WriteLine(line)
            Catch ex As Exception
            End Try
        End If

    End Sub

    ' dgp rev 2/13/07 Scan files based upon current user and current run
    Public Shared Sub CloseLog()

        If mLoggingOn Then
            mLoggingOn = False
            mLogWriter.Close()
        End If

    End Sub

    Public Shared Function LargestDisk() As String

        Dim ms As System.Management.ManagementScope
        Dim oq As System.Management.ObjectQuery
        Dim mos As System.Management.ManagementObjectSearcher
        Dim obj As System.Management.ManagementObject
        Dim ans As Double = 0.0
        Dim name As String

        Dim LargestSpace As Int64
        LargestDisk = ""

        Dim NameHash As New Hashtable
        Dim OrderHash As New Hashtable
        Dim SizeArray As New ArrayList

        ' The "scope" includes the name of the PC and the WMI namespace
        ms = New System.Management.ManagementScope("\\" & FlowServer.ToString & _
         "\root\cimv2")

        ' use WQL to get just the one instance we want.  This should look familiar
        ' to those who are used to SQL
        oq = New System.Management.ObjectQuery("select FreeSpace, Name, VolumeName from Win32_LogicalDisk where DriveType=3")
        ' execute the query
        mos = New System.Management.ManagementObjectSearcher(ms, oq)

        For Each obj In mos.Get
            ans = Convert.ToDouble(obj("FreeSpace")) / (1024.0# * 1024.0#)
            name = Convert.ToString(obj("Name"))
            If ans > LargestSpace Then
                LargestSpace = CInt(ans)
                LargestDisk = name
            End If
        Next

        Return LargestDisk

    End Function

    ' dgp rev 9/13/2010
    Private Shared Function ShareExists(ByVal name) As Boolean

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_Share"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.ToString & _
         "\root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            If management_object("Name").ToString.ToLower = name.ToString.ToLower Then Return True
        Next management_object

        Return False

    End Function

    Public Shared Function VerifyShare() As Boolean


    End Function

    ' dgp rev 9/13/2010
    Public Shared Function CheckShare(ByVal ShareName) As Boolean

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_Share"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.ToString & _
         "\root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            If management_object("Name").ToString.ToLower = ShareName.ToString.ToLower Then Return True
        Next management_object

        Return False

    End Function

    Private Shared mExperShareRootReady = Nothing
    Public Shared ReadOnly Property ExperShareRootReady
        Get
            If mExperShareRootReady IsNot Nothing Then Return mExperShareRootReady

            mExperShareRootReady = CheckShare(ShareExperRoot)
            If Not mExperShareRootReady Then
                Dim dsk = LargestDisk()
                Dim path = String.Format("{0}\{1}", dsk, "Experiments")
                If Not ServerPathExists(path) Then
                    If CreateRemoteFolder(path) Then
                        mExperShareRootReady = CreateShare("Experiments", path)
                    End If
                End If
            End If

            Return mExperShareRootReady

        End Get
    End Property

    Public Shared Function CreateRemoteFolder(ByVal path) As Boolean

        CreateRemoteFolder = False
        Try
            ' assemble the string so the scope represents the remote server
            mPath = String.Format("\\{0}\root\cimv2", mServerName)

            ' connect to WMI on the remote server
            mScope = New ManagementScope(mPath)

            ' create a new instance of the Win32_Share WMI object
            mClass = New ManagementClass("Win32_Process")

            ' set the scope of the new instance to that created above
            mClass.Scope = mScope

            ' Get an input parameters object for this method
            Dim inParams As ManagementBaseObject = mClass.GetMethodParameters("Create")

            ' Fill in input parameter values
            inParams("CommandLine") = String.Format("cmd.exe /c md {0}", path)

            ' Execute the method
            Dim outParams As ManagementBaseObject = mClass.InvokeMethod("Create", inParams, Nothing)


            If CInt(outParams("returnValue")) = 0 Then
                CreateRemoteFolder = True
            End If

        Catch ex As Exception

        End Try

    End Function



    Private Shared mPath As String
    Private Shared mClass As Management.ManagementClass

    Public Shared Function CreateShare(ByVal name, ByVal path) As Boolean

        Try

            ' assemble the string so the scope represents the remote server
            mPath = String.Format("\\{0}\root\cimv2", mServerName)

            ' connect to WMI on the remote server
            mScope = New ManagementScope(mPath)

            ' create a new instance of the Win32_Share WMI object
            mClass = New ManagementClass("Win32_Share")

            ' set the scope of the new instance to that created above
            mClass.Scope = mScope

            ' assemble the arguments to be passed to the Create method
            Dim methodargs = {path, name, "0"}

            ' invoke the Create method to create the share
            Dim results = mClass.InvokeMethod("Create", methodargs)

        Catch ex As Exception

            MsgBox("Oops -- " + ex.Message)
            Return False

        End Try
        Return True

    End Function

    Private mValidShare = Nothing

    ' dgp rev 11/17/08 Flow Root on the Flow Lab Server
    Private Shared mShareExperRoot As String = ""
    Public Shared Property ShareExperRoot() As String
        Get
            If (Not mShareExperRoot = "") Then Return mShareExperRoot
            If (XMLSettings.Exists("ShareExper")) Then
                mShareExperRoot = XMLSettings.GetSetting("ShareExper")
            Else
                mShareExperRoot = "Experiments"
            End If
            Return mShareExperRoot
        End Get
        Set(ByVal value As String)
            mShareExperRoot = value
        End Set
    End Property

    ' dgp rev 2/13/07 Scan files based upon current user and current run
    Public Shared Sub DefineLog(ByVal user As String, ByVal run As String)

        mLoggingOn = False
        Dim path = System.IO.Path.Combine(Server, "Upload")
        path = System.IO.Path.Combine(path, "Logs")
        If Not HelperClasses.Utility.Create_Tree(path) Then Exit Sub
        path = System.IO.Path.Combine(path, user + "_" + run + ".log")
        Try
            mLogWriter = New StreamWriter(path, True)
            mLogWriter.WriteLine(DateTime.Now.ToLongDateString)
            mLogWriter.WriteLine(DateTime.Now.ToLongTimeString)
            mLogWriter.Flush()
        Catch ex As Exception
            Exit Sub
        End Try
        If (Not System.IO.File.Exists(path)) Then Exit Sub
        mLoggingOn = True

    End Sub

    ' dgp rev 2/13/07 Scan files based upon current user and current run
    Public Shared ReadOnly Property UserRunExits(ByVal user, ByVal run) As Boolean
        Get
            Dim path = System.IO.Path.Combine(Server, "Upload")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "FCSRun")
                If (System.IO.Directory.Exists(path)) Then
                    path = System.IO.Path.Combine(path, user)
                    If (System.IO.Directory.Exists(path)) Then
                        path = System.IO.Path.Combine(path, run)
                        If (System.IO.Directory.Exists(path)) Then Return True
                    End If
                End If
            End If
            Return False
        End Get
    End Property


    ' dgp rev 2/13/07 Scan files based upon current user and current run
    Public Shared ReadOnly Property ServerUserRunRoot(ByVal user) As String
        Get
            Dim path = System.IO.Path.Combine(Server, "Upload")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "FCSRun")
                If (System.IO.Directory.Exists(path)) Then
                    path = System.IO.Path.Combine(path, user)
                    If (System.IO.Directory.Exists(path)) Then Return path
                End If
            End If
            Return Nothing
        End Get
    End Property


    ' dgp rev 2/13/07 Scan files based upon current user and current run
    Public Shared ReadOnly Property UserRunPath(ByVal user, ByVal run) As String
        Get
            Dim path = System.IO.Path.Combine(Server, "Upload")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "FCSRun")
                If (System.IO.Directory.Exists(path)) Then
                    path = System.IO.Path.Combine(path, user)
                    If (System.IO.Directory.Exists(path)) Then
                        path = System.IO.Path.Combine(path, run)
                        If (System.IO.Directory.Exists(path)) Then Return path
                    End If
                End If
            End If
            Return Nothing
        End Get
    End Property

    ' dgp rev 2/13/07 download the selected run
    Public Shared Sub Download_Run(ByVal user, ByVal run)

        If (UserRunExits(user, run)) Then

            Dim target_path As String
            target_path = System.IO.Path.Combine(FlowStructure.Data_Root, run)
            If (System.IO.Directory.Exists(target_path)) Then
                If (MsgBox("Overwrite", vbYesNo, "Run Exists") <> vbYes) Then Exit Sub
            Else
                If Not Utility.Create_Tree(target_path) Then
                    MsgBox("Failed to create local run - " + run, MsgBoxStyle.Information)
                    Exit Sub
                End If
            End If

            Dim file
            Dim objFile
            Dim target_spec
            ' loop thru folder
            Dim source_file
            Dim source_path = UserRunPath(user, run)
            For Each file In ServerUserRun(user, run)
                ' create an FCS run object
                source_file = System.IO.Path.Combine(source_path, file)
                objFile = New FCS_Classes.FCS_File(source_file)
                If (objFile.Valid) Then
                    Try
                        ' dgp rev 7/12/07 rename the FCS files to standard format
                        target_spec = System.IO.Path.Combine(target_path, file)
                        System.IO.File.Copy(source_file, target_spec, True)
                        RaiseEvent FSSRemoteEvent(file)
                    Catch ex As Exception
                        RaiseEvent FSSRemoteEvent("Failure: " + file)
                    End Try
                End If
            Next

            ' make instance of run object as target
            Dim objRun = New FCS_Classes.FCSRun(target_path)

            If (objRun.Valid_Run) Then
                FlowStructure.BrowseRun = objRun
            Else
            End If

        Else
            MsgBox(user + " no run found - " + run, MsgBoxStyle.Information)
        End If

    End Sub


    ' dgp rev 2/13/07 Scan files based upon current user and current run
    Public Shared ReadOnly Property ServerUserRun(ByVal user, ByVal run) As ArrayList
        Get
            ServerUserRun = New ArrayList

            Dim path = System.IO.Path.Combine(Server, "Upload")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "FCSRun")
                If (System.IO.Directory.Exists(path)) Then
                    path = System.IO.Path.Combine(path, user)
                    If (System.IO.Directory.Exists(path)) Then
                        path = System.IO.Path.Combine(path, run)
                        If (System.IO.Directory.Exists(path)) Then
                            Dim fcs
                            For Each fcs In System.IO.Directory.GetFiles(path)
                                ServerUserRun.Add(System.IO.Path.GetFileName(fcs))
                            Next
                        End If
                    End If
                End If
            End If
        End Get
    End Property

    ' dgp rev 7/1/09 
    Private Shared mNewServer As String = "NCI-01855598"
    Public Shared ReadOnly Property NewServer As String
        Get
            Return mNewServer
        End Get
    End Property


    Private Shared mServerList() = {mServerName, mNewServer}

    Public Shared Function ServerUserRuns() As ArrayList

        ServerUserRuns = ServerUserRuns(mUsername)

    End Function

    Public Shared Function ServerUserRuns(ByVal user) As ArrayList

        ServerUserRuns = New ArrayList

        Dim path = System.IO.Path.Combine(Server, "Upload")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "FCSRun")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, user)
                If (System.IO.Directory.Exists(path)) Then
                    Dim run
                    For Each run In System.IO.Directory.GetDirectories(path)
                        ServerUserRuns.Add(System.IO.Path.GetFileNameWithoutExtension(run))
                    Next
                End If
            End If
        End If

        path = System.IO.Path.Combine(NewServer, "Upload")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "FCSRun")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, user)
                If (System.IO.Directory.Exists(path)) Then
                    Dim run
                    For Each run In System.IO.Directory.GetDirectories(path)
                        ServerUserRuns.Add(System.IO.Path.GetFileNameWithoutExtension(run))
                    Next
                End If
            End If
        End If

    End Function

    ' dgp rev 7/1/09 
    Public Shared ReadOnly Property Server()
        Get
            Return String.Format("\\{0}", mServerName)
        End Get
    End Property

    ' dgp rev 7/1/09 
    Public Shared Function ServerActualUsers() As ArrayList

        ServerActualUsers = New ArrayList

        Dim path = System.IO.Path.Combine(Server, "Upload")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "FCSRun")
            If (System.IO.Directory.Exists(path)) Then
                If System.IO.Directory.GetDirectories(path).Count = 0 Then
                Else
                    Dim user
                    For Each user In System.IO.Directory.GetDirectories(path)
                        ServerActualUsers.Add(System.IO.Path.GetFileNameWithoutExtension(user))
                    Next
                End If
            End If
        End If

    End Function

    ' dgp rev 7/1/09 
    Public Shared Function ServerUserDataExists() As Boolean

        ServerUserDataExists = False

        Dim path = System.IO.Path.Combine(Server, "Upload")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "FCSRun")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, mUsername)
                If (System.IO.Directory.Exists(path)) Then ServerUserDataExists = (System.IO.Directory.GetDirectories(path).Length > 0)
            End If
        End If

    End Function

    Shared Function ServerDataFindRun(RunName As Object) As Object
        Throw New NotImplementedException
    End Function

    Shared Function ServerUserRunPath(p1 As String, RunName As Object) As Object
        Throw New NotImplementedException
    End Function

End Class
