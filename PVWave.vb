' Name:     PV-Wave Distribution Class
' Author:   Donald G Plugge
' Date:     2/22/07 
' Purpose:  Class to handle the configuration of PV-Wave CPR files

Imports Microsoft.Win32
Imports System.IO
Imports HelperClasses
Imports System.Xml
Imports FCS_Classes
Imports System.Management
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class PVWave

    Public Delegate Sub ExitedEventHandler(ByVal SomeString As String)
    Public Event ExitedEvent As ExitedEventHandler

    Private mLogger As Logger
    Public Property Logger() As Logger
        Get
            Return mLogger
        End Get
        Set(ByVal value As Logger)
            mLogger = value
        End Set
    End Property

    Private Sub Log_Info(ByVal text As String)

        If (mLogger Is Nothing) Then Exit Sub

        mLogger.Log_Info(text)

    End Sub


    ' dgp rev 3/27/08 login information incase of priviledge issues
    Private mwmiOptions As New ConnectionOptions
    Public Property wmiOptions() As ConnectionOptions
        Get
            Return mwmiOptions
        End Get
        Set(ByVal value As ConnectionOptions)
            mwmiOptions = value
        End Set
    End Property
    ' dgp rev 3/27/08
    Private mImpFlag As Boolean = False
    Public Property ImpFlag() As Boolean
        Get
            Return mImpFlag
        End Get
        Set(ByVal value As Boolean)
            mImpFlag = value
        End Set
    End Property
    ' dgp rev 3/27/08
    Private mFlagServer As Boolean = False
    Public Property FlagServer() As Boolean
        Get
            Return mFlagServer
        End Get
        Set(ByVal value As Boolean)
            mFlagServer = value
        End Set
    End Property
    ' dgp rev 3/27/08
    Private mAccount As String
    Public Property Account() As String
        Get
            Return mAccount
        End Get
        Set(ByVal value As String)
            mAccount = value
        End Set
    End Property
    ' dgp rev 3/27/08
    Private mPassword As String
    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property
    ' dgp rev 3/27/08
    Private mObjImp As New RunAs_Impersonator
    Public Property ObjImp() As RunAs_Impersonator
        Get
            Return mObjImp
        End Get
        Set(ByVal value As RunAs_Impersonator)
            mObjImp = value
        End Set
    End Property

    Private mKey As String
    Private mTest As String

    Private mSep As Char = System.IO.Path.DirectorySeparatorChar

    Private mdyn As New Dynamic("PVWave")
    Public Property pvwXML() As Dynamic
        Get
            Return mdyn
        End Get
        Set(ByVal value As Dynamic)
            mdyn = value
        End Set
    End Property

    Shared Shell_List As New Collection
    Shared Wave_List As New Collection
    Shared cmd_path As String
    Shared cmd_name As String = "my_cmd"
    Shared common_cmd As String = "C:\Windows\System32\CMD.exe"
    Public Shared Username As String
    Shared drvlst As Collection
    Shared PVW_WorkOld As String
    Public Shared PVW_Browse_Run As FCS_Classes.FCSRun

    Private mLocalVersion As String


    Public Sub UpgradeTest()

        Dim test As String = DistributionServer.TestVersion
        Server_Ver = test
        Dim source As String = "\\" + Server + "\" + Share + "\Versions\" + test
        Dim target As String = System.IO.Path.Combine(FlowStructure.Dist_Root.ToString, test)

        Dim UpdateFlag As Boolean = False
        Dim SuccessFlag As Boolean = False

        If (ImpFlag) Then ObjImp.ImpersonateStart("NIH", Account, Password)
        If (System.IO.Directory.Exists(source)) Then
            If (Not System.IO.Directory.Exists(target)) Then
                UpdateFlag = True
                Err.Clear()
                Utility.DirectoryCopy(source, target)
                SuccessFlag = (Err.Number = 0)
            End If
        End If
        If (ImpFlag) Then ObjImp.ImpersonateStop()

        If (UpdateFlag And Not SuccessFlag) Then MsgBox("Update Error " + Err.Description, MsgBoxStyle.Information)

        If (UpdateFlag And SuccessFlag) Then
            ' dgp rev 2/6/08 make the new version active
            m_Local_List = Nothing
        End If

    End Sub

    Public Sub ReSetLocalList()

        m_Local_List = Nothing

    End Sub

    Private mOperatorCheck As Boolean = False
    Private mOperatorFlag As Boolean
    Public ReadOnly Property OperatorFlag As Boolean
        Get
            If Not mOperatorCheck Then
                mOperatorFlag = FlowServer.Check_Admin
            End If
            Return mOperatorFlag
        End Get
    End Property

    Private mTestMode As Boolean = False

    Public Property TestMode As Boolean
        Get
            If (pvwXML.Exists("Mode")) Then
                ' either from the XML file or...
            Else
                pvwXML.PutSetting("Mode", "Production")
            End If
            Return (pvwXML.GetSetting("Mode") = "Test")
        End Get
        Set(ByVal value As Boolean)
            If value Then
                pvwXML.PutSetting("Mode", "Test")
            Else
                pvwXML.PutSetting("Mode", "Production")
            End If
        End Set
    End Property

    ' dgp rev 5/12/2011 keeping track of the version
    Public Sub Upgrade(ByVal DesiredVersion)

        Server_Ver = DesiredVersion
        Dim source As String = "\\" + Server + "\" + Share + "\Versions\" + Server_Ver
        Dim target As String = System.IO.Path.Combine(FlowStructure.Dist_Root.ToString, Server_Ver)

        Dim UpdateFlag As Boolean = False
        Dim SuccessFlag As Boolean = False

        If (ImpFlag) Then ObjImp.ImpersonateStart("NIH", Account, Password)
        If (System.IO.Directory.Exists(source)) Then
            If (Not System.IO.Directory.Exists(target)) Then
                Err.Clear()
                Utility.DirectoryCopy(source, target)
                SuccessFlag = (Err.Number = 0)
            Else
                SuccessFlag = True
            End If
            UpdateFlag = True
        End If
        If (ImpFlag) Then ObjImp.ImpersonateStop()

        If (UpdateFlag And Not SuccessFlag) Then MsgBox("Update Error " + Err.Description, MsgBoxStyle.Information)

        If (UpdateFlag And SuccessFlag) Then
            ' dgp rev 2/6/08 make the new version active
            m_Local_List = Nothing
        End If

    End Sub

    ' dgp rev 5/12/2011 keeping track of the version
    Private mPVWaveInstalled = Nothing
    Public ReadOnly Property PVWaveInstalled() As Boolean
        Get
            If mPVWaveInstalled Is Nothing Then Get_PVWave_Path()
            Return mPVWaveInstalled
        End Get
    End Property

    ' dgp rev 7/18/07 Server and Share Info
    Private mPVW_Server As String
    Public Property PVW_Server() As String
        Get
            Return mPVW_Server
        End Get
        Set(ByVal value As String)
            mPVW_Server = value
        End Set
    End Property

    Private mPVW_Server_IP As String
    Public Property PVW_Server_IP() As String
        Get
            Return mPVW_Server_IP
        End Get
        Set(ByVal value As String)
            mPVW_Server_IP = value
        End Set
    End Property

    Private mPVW_Dist_Share As String
    Public Property PVW_Dist_Share() As String
        Get
            Return mPVW_Dist_Share
        End Get
        Set(ByVal value As String)
            mPVW_Dist_Share = value
        End Set
    End Property

    Private mPVW_Data_Share As String
    Public Property PVW_Data_Share() As String
        Get
            Return mPVW_Data_Share
        End Get
        Set(ByVal value As String)
            mPVW_Data_Share = value
        End Set
    End Property

    Private mPVW_Work_Share As String
    Public Property PVW_Work_Share() As String
        Get
            Return mPVW_Work_Share
        End Get
        Set(ByVal value As String)
            mPVW_Work_Share = value
        End Set
    End Property

    Public proc_list As Collection
    Public hndl_list As Collection

    Public cmd_str As String
    Private mParams As String = " /K wave -r flow_control"
    Public output_str As String

    Public proc_name As String
    Public proc_id As Integer

    ' dgp rev 7/18/07 data root assigned
    Private mChkDataRoot As Boolean = False

    ' dgp rev 11/28/07 get the DataRoot from one place
    ' dgp rev 11/28/07 how is data root assigned - scan or use FCS Files list
    Private mDataRoot As String
    Public Property DataRoot() As String
        Get
            If (mDataRoot Is Nothing) Then mDataRoot = FlowStructure.Data_Root
            Return mDataRoot
        End Get
        Set(ByVal value As String)
            mDataRoot = value
        End Set
    End Property

    Private mSession As String
    Public Property Session() As String
        Get
            Return mSession
        End Get
        Set(ByVal value As String)
            mSession = value
        End Set
    End Property

    ' dgp rev 5/12/09 Parse project from Cur Work
    ' humm
    Private Function Parse_Project() As String

        Return ""

    End Function

    Private mProject
    Public Property Project() As String
        Get
            If mProject Is Nothing Then mProject = Parse_Project()
            Return mProject
        End Get
        Set(ByVal value As String)
            mProject = value
        End Set
    End Property

    ' dgp rev 5/29/08 Work Validation
    Private Function Check_Work(ByVal path) As Boolean

        Check_Work = False
        ' dgp rev 9/29/2010 Use FCS Files list object
        Dim objFCSList As FCS_List
        ' dgp rev 5/29/08 is work path valid
        If (System.IO.Directory.Exists(path)) Then
            objFCSList = New FCS_List(path)
            If (System.IO.File.Exists(objFCSList.List_Spec)) Then
                Dim sr As New StreamReader(objFCSList.List_Spec)
                While (Not sr.EndOfStream)
                    If (System.IO.File.Exists(sr.ReadLine)) Then
                        Check_Work = True
                        Exit While
                    End If
                End While
                sr.Close()
            End If
        End If

    End Function

    ' dgp rev 5/23/08 what the heck is going on here -- curwork = curwork?
    Private mValidWorkPath As Boolean = False
    Public ReadOnly Property ValidWorkPath() As Boolean
        Get
            If (FlowStructure.CurWork = "") Then Return False
            Return System.IO.Directory.Exists(FlowStructure.CurWork)
        End Get
    End Property
    ' dgp rev 7/18/07 path to current work
    ' dgp rev 2/27/08 most FlowRoot info contained in FlowStructure.  Dynamic info kept here.
    ' dgp rev 5/28/08 where should the current work be validated.  Here?
    Private mWorkRoot
    Public Property Work_Root() As String
        Get
            If mWorkRoot Is Nothing Then mWorkRoot = FlowStructure.Work_Root
            Return mWorkRoot
        End Get
        Set(ByVal value As String)
            If System.IO.Directory.Exists(value) Then mWorkRoot = value
        End Set
    End Property

    ' dgp rev 6/15/09 After the process is complete, save latest work
    Public Sub SaveCurWork()

        Dim path
        If (FlowServer.Server_Up()) Then
            If (FlowServer.UserExists) Then
                If (System.IO.Directory.Exists(FlowServer.UserPath)) Then
                    Dim Source = FlowStructure.CurWork
                    Dim Target = FlowServer.UserPath
                    Dim arr = Split(FlowStructure.CurWork, System.IO.Path.DirectorySeparatorChar)
                    If arr.Length > 3 Then
                        path = FlowServer.UserPath
                        Dim idx
                        For idx = 3 To 1 Step -1
                            path = System.IO.Path.Combine(path, arr(arr.Length - idx))
                        Next
                        If (Not System.IO.Directory.Exists(path)) Then
                            If (Utility.Create_Tree(path)) Then
                                Utility.DirectoryCopy(Source, path)
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private mPVWCmd As Process
    Public Property PVWCmd() As Process
        Get
            Return mPVWCmd
        End Get
        Set(ByVal value As Process)
            mPVWCmd = value
        End Set
    End Property

    Private mAutoStart As Boolean = True
    Public Property AutoStart() As Boolean
        Get
            Return mAutoStart
        End Get
        Set(ByVal value As Boolean)
            mAutoStart = value
        End Set
    End Property

    ' dgp rev 5/10/07 
    Private Shared Function PVW_from_CFG() As Hashtable

        PVW_from_CFG = New Hashtable
        ' dgp rev 11/1/06 read the last configuration file in text format
        ' dgp rev 11/3/06 add a settings hash

        Dim line As String
        Dim fig_arr() As String

        Dim txt_cfg_file As String = ""

        ' dgp rev 7/18/07 the cfg file is for backward compatibility
        If (CFG_File Is Nothing) Then Return PVW_from_CFG

        '        Log_Info("Checking for old config...")
        '        Log_Info(CFG_File.ToString)
        If (System.IO.File.Exists(CFG_File)) Then
            '            Log_Info("Reading text " + CFG_File)
            Dim sr As New System.IO.StreamReader(CFG_File)
            PVW_from_CFG = New Hashtable
            While (Not sr.EndOfStream)
                line = sr.ReadLine()
                '                Log_Info(line)
                fig_arr = Split(line, Mid(line, 1, 1))
                PVW_from_CFG.Add(fig_arr(1), fig_arr(2))
                '                Log_Info(fig_arr(1) + " => " + fig_arr(2))
            End While
            sr.Close()
        Else
            '            Log_Info("No old config")
        End If

    End Function

    ' dgp rev 7/18/07 get the XML config file 
    Private Shared Function PVW_from_XML() As Hashtable

        PVW_from_XML = New Hashtable

        ' dgp rev 5/1/07 Dynamic settings
        '        dyn = New Dynamic(system.io.path.combine(user_path, "Dynamic.XML"))
        If (System.IO.File.Exists(FlowStructure.PVW_XML_File)) Then
            ' dgp rev 3/31/2011 add filename key to attribute list
            '            If Not PVW_from_XML.ContainsKey("Filename") Then PVW_from_XML.Add("Filename", FlowStructure.PVW_XML_File)
            '            Log_Info("Reading XML " + FlowStructure.PVW_XML_File)
            PVW_XML_Doc = New XmlDocument
            Dim XMLNode As XmlNode
            PVW_XML_Doc.Load(FlowStructure.PVW_XML_File)
            XMLNode = PVW_XML_Doc.SelectSingleNode("Environment/FlowControl/Settings")
            If (XMLNode Is Nothing) Then
                '                PVW_from_XML.Add("Error", True)
            Else
                Dim idx
                For idx = 0 To XMLNode.Attributes.Count - 1
                    PVW_from_XML.Add(XMLNode.Attributes.Item(idx).Name, XMLNode.Attributes.Item(idx).Value)
                    '                    Log_Info(XMLNode.Attributes.Item(idx).Name + " => " + XMLNode.Attributes.Item(idx).Value)
                Next
            End If

        End If

    End Function

    ' dgp rev 5/23/07 Fill in the defaults for work_path and setup_path
    Private Function PVW_Defaults() As Hashtable

        PVW_Defaults = New Hashtable
        PVW_Defaults.Add("setup_path", FlowStructure.Settings)
        PVW_Defaults.Add("work_path", FlowStructure.CurWork)

    End Function
    ' dgp rev 5/12/09 Load the memory with presistant values
    Private Shared Sub Load_PVW_Attr()

        If (mBetaFlag) Then
            mPVW_Attr = PVW_from_CFG()
        Else
            mPVW_Attr = PVW_from_XML()
        End If

        If (mPVW_Attr.Contains("work_path")) Then
            FlowStructure.CurWork = mPVW_Attr.Item("work_path")
        End If

        If (mPVW_Attr.Contains("setup_path")) Then
            FlowStructure.Settings = mPVW_Attr.Item("setup_path")
        End If

    End Sub
    ' dgp rev 7/17/07 PVWave uses an XML file to hold settings
    ' dgp rev 7/17/07 XML file name, XML Document and XML Attributes 
    Private Shared mPVW_Attr As Hashtable = Nothing
    Public Shared Property PVW_Attr() As Hashtable
        Get
            If mPVW_Attr Is Nothing Then
                mPVW_Attr = PVW_from_XML()
                If mPVW_Attr.Count = 0 Then mPVW_Attr = PVW_from_CFG()
            End If
            Return mPVW_Attr
        End Get
        Set(ByVal value As Hashtable)
            mPVW_Attr = value
        End Set
    End Property

    Private Shared mPVW_XML_Doc As XmlDocument
    Public Shared Property PVW_XML_Doc() As XmlDocument
        Get
            Return mPVW_XML_Doc
        End Get
        Set(ByVal value As XmlDocument)
            mPVW_XML_Doc = value
        End Set
    End Property

    ' dgp rev 4/19/07 server scan flag
    Private m_scan_on_set As Boolean = False
    Public Property Scan_On_Set() As Boolean
        Get
            Return m_scan_on_set
        End Get
        Set(ByVal value As Boolean)
            m_scan_on_set = value
        End Set
    End Property

    ' dgp rev 2/22/07 member for local root
    Private m_Server_Valid As Boolean = False
    Public Property Server_Valid() As Boolean
        Get
            Return m_Server_Valid
        End Get
        Set(ByVal value As Boolean)
            m_Server_Valid = value
        End Set
    End Property

    ' dgp rev 2/22/07 scan for local Servers
    Public Sub Scan_Server_Vers()

        Dim Root_Path As String = "\\" + Server + "\" + Share + "\Versions"
        Dim tmp As New Collection

        If (ImpFlag) Then ObjImp.ImpersonateStart("NIH", Account, Password)
        If (System.IO.Directory.Exists(Root_Path)) Then
            Dim item
            For Each item In System.IO.Directory.GetDirectories(Root_Path)
                tmp.Add(System.IO.Path.GetFileName(item))
            Next
        End If
        If (ImpFlag) Then ObjImp.ImpersonateStop()

        Server_List = tmp

    End Sub

    Private mRemotePath As String = ""
    Private ReadOnly Property RemotePath As String
        Get
            ' dgp rev 5/18/2011 make protocol path dynamic based upon selected user
            mRemotePath = String.Format("\\{0}\{1}\Users\{2}", FlowServer.FlowServer, FlowServer.ShareFlow, Username)
            mRemotePath = System.IO.Path.Combine(mRemotePath, "Logging")
            Return mRemotePath
        End Get
    End Property

    Private mUploadFile As String
    Public ReadOnly Property UploadFile As String
        Get
            Dim name = String.Format("PVW_{0}.log", Format(Now(), "yyMMddhhmm"))
            mUploadFile = System.IO.Path.Combine(RemotePath, name)
            Return mUploadFile
        End Get
    End Property

    Private mOneTimeSetup As Boolean = False
    Private mLoggable As Boolean = False
    Private mLogSW As StreamWriter

    ' dgp rev 8/10/2011 
    Public Sub SystemLogAppend(ByVal txt As String)

        If Not mOneTimeSetup Then
            mOneTimeSetup = True
            If Utility.Create_Tree(RemotePath) Then
                Try
                    mLogSW = New StreamWriter(UploadFile)
                    mLogSW.WriteLine(UploadFile)
                    mLogSW.WriteLine(FlowStructure.CurDist)
                    mLogSW.WriteLine(txt)
                    mLogSW.Close()
                    mLoggable = True
                Catch ex As Exception
                    mLoggable = False
                End Try
            End If
        Else
            If mLoggable Then
                Try
                    mLogSW = New StreamWriter(UploadFile)
                    mLogSW.WriteLine(txt)
                    mLogSW.Close()
                Catch ex As Exception
                    mLoggable = False
                End Try
            End If
        End If

    End Sub

    ' dgp rev 8/31/06 create a CFG settings file
    Public Function Write_CFG(ByVal spec As String)

        Write_CFG = False

        ' dgp rev 11/1/06 only add work_path and setup_path to xml file
        Dim ie As IDictionaryEnumerator = PVW_Attr.GetEnumerator()

        Try
            Dim sw As New System.IO.StreamWriter(spec, False)
            Dim sep As String = "@"
            While ie.MoveNext()
                sw.WriteLine(sep + ie.Key + sep + ie.Value + sep)
            End While
            sw.Close()
            Write_CFG = True
        Catch ex As Exception

        End Try

    End Function

    ' dgp rev 8/31/06 create an XML settings file
    Public Function Write_CFG()

        Write_CFG = False

        If (CFG_File Is Nothing) Then Return Write_CFG

        Dim path As String = System.IO.Path.GetDirectoryName(CFG_File)
        If (Not System.IO.Directory.Exists(path)) Then Return Write_CFG

        ' dgp rev 11/1/06 only add work_path and setup_path to xml file
        Dim ie As IDictionaryEnumerator = PVW_Attr.GetEnumerator()

        Dim sw As New System.IO.StreamWriter(CFG_File, False)
        Dim sep As String = "@"
        While ie.MoveNext()
            sw.WriteLine(sep + ie.Key + sep + ie.Value + sep)
        End While
        sw.Close()
        Write_CFG = True

    End Function


    ' dgp rev 7/17/07 PVW Sessions needs an XML config file
    Public Sub Create_Configs()

        If (PVW_Attr Is Nothing) Then
            mPVW_Attr = PVW_Defaults()
        Else
            If (mPVW_Attr.Contains("work_path")) Then
                mPVW_Attr.Item("work_path") = FlowStructure.CurWork
            Else
                mPVW_Attr.Add("work_path", FlowStructure.CurWork)
            End If
            If (mPVW_Attr.Contains("setup_path")) Then
                mPVW_Attr.Item("setup_path") = FlowStructure.Settings
            Else
                mPVW_Attr.Add("setup_path", FlowStructure.Settings)
            End If
        End If

        If (mBetaFlag) Then
            Write_CFG(CFG_New)
            Write_CFG(CFG_Old)
            Write_XML()
        Else
            Write_CFG(CFG_New)
            Write_CFG(CFG_Old)
            Write_XML()
        End If

    End Sub

    ' dgp rev 2/22/07 scan for local versions
    Public Sub Scan_Local_Vers()
        m_Local_List = New ArrayList
        If (System.IO.Directory.Exists(FlowStructure.Dist_Root)) Then
            Dim item
            For Each item In System.IO.Directory.GetDirectories(FlowStructure.Dist_Root)
                m_Local_List.Add(System.IO.Path.GetFileName(item))
            Next
        End If
    End Sub

    ' dgp rev 7/18/07
    Private Shared m_CFG_File As String

    ' dgp rev 1/21/2011
    Public ReadOnly Property CFG_Old As String
        Get
            Return System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.FlowRoot, "Users"), Username + ".cfg")
        End Get
    End Property

    ' dgp rev 1/21/2011
    Public ReadOnly Property CFG_New As String
        Get
            Return System.IO.Path.Combine(FlowStructure.Settings, Username + ".cfg")
        End Get
    End Property

    Public Shared ReadOnly Property CFG_File() As String
        Get
            If (mBetaFlag) Then
                '                Log_Info("Beta CFG File")
                m_CFG_File = System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.FlowRoot, "Users"), Username + ".cfg")
            Else
                If (FlowStructure.Settings IsNot Nothing) Then
                    m_CFG_File = System.IO.Path.Combine(FlowStructure.Settings, Username + ".cfg")
                Else
                    m_CFG_File = System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.FlowRoot, "Users"), Username + ".cfg")
                End If
            End If
            Return m_CFG_File
        End Get
    End Property
    ' dgp rev 5/23/07 
    Private m_Setup_Path As String
    Public ReadOnly Property Setup_Path() As String
        Get
            If m_Setup_Path Is Nothing Then m_Setup_Path = FlowStructure.PersonalSettingsRoot
            Return m_Setup_Path
        End Get
    End Property

    ' dgp rev 2/23/07 member for local root
    Private m_Share As String
    Public Property Share() As String
        Get
            Return m_Share
        End Get
        Set(ByVal value As String)
            ' did anything change
            If (m_Share <> value) Then
                m_Share = value
                ' when Share root changed, rescan
                If (Scan_On_Set And Not Server Is Nothing) Then
                    Scan_Server_Vers()
                    Validate_Server()
                End If
            End If
        End Set
    End Property


    ' dgp rev 2/23/07 member for local root
    Private m_Server As String
    Public Property Server() As String
        Get
            Return m_Server
        End Get
        Set(ByVal value As String)
            ' did anything change
            If (m_Server <> value) Then
                m_Server = value
                ' when Server root changed, rescan
                'Scan_Server_Vers()
                'Validate_Server()
            End If
        End Set
    End Property

    ' dgp rev 2/22/07 list of local versions
    Private m_Local_List As ArrayList = Nothing
    Public Property Local_List() As ArrayList
        Get
            If m_Local_List Is Nothing Then Scan_Local_Vers()
            Return m_Local_List
        End Get
        Set(ByVal value As ArrayList)
            m_Local_List = value
        End Set
    End Property
    ' dgp rev 2/22/07 list of local versions
    Private m_Server_List As Collection
    Public Property Server_List() As Collection
        Get
            Return m_Server_List
        End Get
        Set(ByVal value As Collection)
            m_Server_List = value
        End Set
    End Property
    ' dgp rev 2/22/07 list of local versions
    Private m_Log_File As String
    Public Property Log_File() As String
        Get
            Return m_Log_File
        End Get
        Set(ByVal value As String)
            m_Log_File = value
        End Set
    End Property

    ' dgp rev 3/27/08 Verify connection to server
    Public Function Check_Server() As Boolean

        Check_Server = True

        Dim mCIMV2 As String = "\\" & Server & "\root\cimv2:Win32_GroupUser"
        Dim mClass As ManagementClass = New ManagementClass("Win32_GroupUser")
        Dim myScope = New ManagementScope(mCIMV2, wmiOptions)

        '* connect to WMI namespace
        Try
            myScope.Connect()
            FlagServer = True
        Catch ex As Exception
            Check_Server = False
        End Try

    End Function


    ' dgp rev 2/23/07 Validate the current Server path
    Public Sub Validate_Server()

        Server_Valid = False

        If (Server Is Nothing) Then Exit Sub
        If (Share Is Nothing) Then Exit Sub
        If (Server_Ver Is Nothing) Then Exit Sub

        Dim fullpath As String = "\\" + Server + "\" + Share + "\Versions"

        fullpath = System.IO.Path.Combine(fullpath, Server_Ver)
        If (ImpFlag) Then ObjImp.ImpersonateStart("NIH", Account, Password)
        If (System.IO.Directory.Exists(fullpath)) Then
            Server_Valid = (System.IO.File.Exists(System.IO.Path.Combine(fullpath, "flow_control.cpr")))
        End If
        If (ImpFlag) Then ObjImp.ImpersonateStop()

    End Sub

    ' dgp rev 2/22/07 member for local root
    ' dgp rev 2/22/07 member for local version
    ' dgp rev 2/26/08 Local version is persistent in the XML file
    ' 1) check for fast variable
    ' 2) 
    Private mLocalVerFlag As Boolean = False
    '    Private mLocalVerOnce As Boolean = False
    '    Private mLocalVerFast As Boolean = False

    Public Property Local_Valid() As Boolean
        Get
            Return mLocalVerFlag
        End Get
        Set(ByVal value As Boolean)
            mLocalVerFlag = value
        End Set
    End Property

    ' dgp rev 2/23/07 Validate the current distribution path
    ' dgp rev 2/26/08 Folder must exist with Flow_Control.CPR file
    Public Function Validate(ByVal test As String) As Boolean

        Validate = False

        If FlowStructure.NoDist Then Return False

        If (System.IO.Directory.Exists(FlowStructure.FlowRoot)) Then

            Dim fullpath As String
            fullpath = System.IO.Path.Combine(FlowStructure.Dist_Root, test)
            If (System.IO.Directory.Exists(fullpath)) Then
                Validate = (System.IO.File.Exists(System.IO.Path.Combine(fullpath, "flow_control.cpr")))
            End If
        End If

        Validate = (Validate And PVWaveInstalled)

        Return Validate

    End Function
    ' dgp rev 5/7/09 Beta version Flag, uses the old startup sequence
    Private Shared mBetaFlag As Boolean

    Public Shared ReadOnly Property BetaFlag As Boolean
        Get
            If (FlowStructure.CurDist IsNot Nothing) Then Return False
            If (Not mBetaFlag = FlowStructure.CurDist.ToString.ToLower.Contains("beta")) Then
                mBetaFlag = FlowStructure.CurDist.ToString.ToLower.Contains("beta")
                Load_PVW_Attr()
            End If
            Return mBetaFlag
        End Get
    End Property



    Private Function DownloadVersion(ByVal version) As Boolean

        Return False

    End Function

    ' dgp  rev 9/20/2011 Initialize local version
    Private Sub InitLocalVer()

        If (FlowStructure.CurDist IsNot Nothing) Then
            mBetaFlag = FlowStructure.CurDist.ToString.ToLower.Contains("beta")
            Load_PVW_Attr()
        End If

    End Sub

    ' dgp rev 2/22/07 member for local Server
    Private m_Server_Ver As String
    Public Property Server_Ver() As String
        Get
            Return m_Server_Ver
        End Get
        Set(ByVal value As String)
            m_Server_Ver = value
            Validate_Server()
        End Set
    End Property

    ' dgp rev 2/22/07 member for local root

    ' dgp rev 5/10/07 
    Private mPrevSetup

    ' dgp rev 5/29/08 Work Validation
    Private Function Check_Last(ByVal path) As Boolean

        Check_Last = False

        If (Not System.IO.Directory.Exists(path)) Then Exit Function
        Dim objFCSList As FCS_List = New FCS_List(path)

        If (Not System.IO.File.Exists(objFCSList.List_Spec)) Then Exit Function

        Dim sr As New StreamReader(objFCSList.List_Spec)
        Dim line As String
        While (Not sr.EndOfStream)
            line = sr.ReadLine
            If (System.IO.File.Exists(line)) Then
                Log_Info("Data")
                Log_Info(line)
                Check_Last = True
                Exit While
            Else
                Log_Info("No Data Exists")
                Log_Info(line)
            End If
        End While
        sr.Close()

    End Function

    ' dgp rev 5/28/08 Select most recent data run

    ' dgp rev 5/28/08 Create a new run based upon most recent data

    ' dgp rev 5/15/07
    Private Sub Convert_Work()

        Dim root_path As String = System.IO.Path.GetDirectoryName(PVW_WorkOld)
        Dim new_path As String = root_path + "\Users\" + Username + "\Work"

        Dim subfld As String

        MsgBox("New Work Area" + vbCrLf + new_path, MsgBoxStyle.Information, "Updating Work Area")

        Utility.Create_Tree(new_path)

        Dim fld
        For Each fld In System.IO.Directory.GetDirectories(PVW_WorkOld)
            Try
                subfld = System.IO.Path.Combine(new_path, System.IO.Path.GetFileName(fld))
                Utility.DirectoryCopy(System.IO.Path.GetFileNameWithoutExtension(fld), subfld)
            Catch ex As Exception

            End Try
        Next

        '        FlowStructure.Work_Root = new_path

    End Sub

    ' dgp rev 5/10/07 
    Public Function PVW_Data_List() As Boolean

        PVW_Data_List = False
        If (Not DataRoot Is Nothing) Then
            PVW_Data_List = (System.IO.Directory.GetDirectories(DataRoot).Length > 0)
        End If

    End Function


    ' dgp rev 5/10/07 
    ' dgp rev 12/9/07 Change Data Path in current session
    Public Sub PVW_Change_Data(ByVal Data_Path As String)

        ' Log_Info("Create Data List")

        Dim Run As New FCS_Classes.FCSRun(Data_Path)
        If (Run.Valid_Run) Then
            Log_Info(Run.Data_Path)
            ' dgp rev 9/29/2010 Use the FCS Files list object
            Dim objFCSList As FCS_List = New FCS_List(FlowStructure.CurWork)
            Dim sw As New StreamWriter(objFCSList.List_Spec)
            Dim idx

            For idx = 0 To Run.FCS_cnt - 1
                Log_Info(Run.FCS_Files(idx).FullSpec)
                sw.WriteLine(Run.FCS_Files(idx).FullSpec)
            Next
            sw.Close()
        End If

    End Sub


    ' dgp rev 5/10/07 
    Public Sub PVW_Create_DataList(ByVal New_Work As String)

        Dim Data_Path As String = ""

        ' Log_Info("Create Data List")

        If (PVW_Browse_Run.Valid_Run) Then
            Log_Info(PVW_Browse_Run.Data_Path)
            ' dgp rev 9/29/2010 Use the FCS Files list object
            Dim objFCSList As FCS_List = New FCS_List(New_Work)
            Dim sw As New StreamWriter(objFCSList.List_Spec)
            Dim idx

            For idx = 0 To PVW_Browse_Run.FCS_cnt - 1
                Log_Info(PVW_Browse_Run.FCS_Files(idx).FullSpec)
                sw.WriteLine(PVW_Browse_Run.FCS_Files(idx).FullSpec)
            Next
            sw.Close()
        End If

    End Sub

    ' dgp rev 5/15/07
    Public Sub Init_Work()

    End Sub

    Private Sub Get_Dynamic()

        Dim key As String
        key = "Server"
        If (pvwXML.Exists(key)) Then
            Server = pvwXML.GetSetting(key)
        Else
            Server = "nt-eib-10-6b16"
            pvwXML.PutSetting(key, Server)
        End If
        key = "DistShare"
        If (pvwXML.Exists(key)) Then
            Share = pvwXML.GetSetting(key)
        Else
            Share = "Distribution"
            pvwXML.PutSetting(key, Share)
        End If

    End Sub

    Private mPVWavePath = Nothing

    ' dgp rev 6/24/09 PVW Path
    Public Property PVWPath() As String
        Get
            If (PVWaveInstalled) Then Return mPVWavePath
            Return ""
        End Get
        Set(ByVal value As String)
            If (System.IO.File.Exists(value)) Then
                mPVWavePath = value
                mPVWaveInstalled = True
                Dim path As String = System.Environment.GetEnvironmentVariable("path")
                path = path + ";" + value
                System.Environment.SetEnvironmentVariable("path", path)
            End If
        End Set
    End Property

    Private mPVWaveRootPath As String
    Public ReadOnly Property PVWaveRootPath As String
        Get
            Return mPVWaveRootPath
        End Get
    End Property

    ' dgp rev 6/24/09 Check registry for PV-Wave path
    Private Function PVWave_Registry() As Boolean

        Dim RegKey As RegistryKey
        Dim KeyList() As String
        mPVWaveInstalled = False
        PVWave_Registry = False

        Dim RegPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WAVE.EXE"

        Try
            RegKey = Registry.LocalMachine.OpenSubKey(RegPath, False)
            If (RegKey IsNot Nothing) Then
                KeyList = RegKey.GetValueNames
                If (RegKey.ValueCount > 0) Then
                    Dim filename As String = RegKey.GetValue("")
                    mPVWavePath = System.IO.Path.GetDirectoryName(filename).ToString
                    mPVWaveInstalled = True
                    mPVWaveRootPath = RegKey.GetValue("Path")
                    PVWave_Registry = True
                End If
            End If
        Catch ex As Exception

        End Try

    End Function

    ' dgp rev 6/24/09 Check registry for PV-Wave path
    Private Function PVWave_RegistryII() As Boolean

        Dim RegKey As RegistryKey
        Dim KeyList() As String
        mPVWaveInstalled = False
        PVWave_RegistryII = False

        Dim RegPath As String = "SOFTWARE\Visual Numerics\PV-WAVE"

        Try
            RegKey = Registry.LocalMachine.OpenSubKey(RegPath, False)
            If (RegKey IsNot Nothing) Then
                KeyList = RegKey.GetSubKeyNames
                If (RegKey.SubKeyCount > 0) Then
                    RegPath = RegPath + "\" + KeyList.GetValue(KeyList.Length - 1)
                    RegKey = Registry.LocalMachine.OpenSubKey(RegPath, False)
                    RegPath = RegPath + "\" + "Environment"
                    RegKey = Registry.LocalMachine.OpenSubKey(RegPath, False)
                    Dim path = RegKey.GetValue("WAVE_DIR")
                    mPVWaveRootPath = RegKey.GetValue("VNI_DIR")
                    If System.IO.Directory.Exists(path) Then
                        path = System.IO.Path.Combine(path, "bin")
                        If System.IO.Directory.Exists(path) Then
                            If System.IO.Directory.Exists(System.IO.Path.Combine(path, "bin.i386nt")) Then
                                path = System.IO.Path.Combine(path, "bin.i386nt")
                                If System.IO.Directory.Exists(path) Then
                                    Dim filename = System.IO.Path.Combine(path, "wave.exe")
                                    If System.IO.File.Exists(filename) Then
                                        mPVWavePath = path
                                        mPVWaveInstalled = True
                                        PVWave_RegistryII = True
                                    End If
                                End If
                            End If
                            If System.IO.Directory.Exists(System.IO.Path.Combine(path, "bin.win64")) Then
                                path = System.IO.Path.Combine(path, "bin.win64")
                                If System.IO.Directory.Exists(path) Then
                                    Dim filename = System.IO.Path.Combine(path, "wave.exe")
                                    If System.IO.File.Exists(filename) Then
                                        mPVWavePath = path
                                        mPVWaveInstalled = True
                                        PVWave_RegistryII = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Function

    ' dgp rev 2/28/07
    Private Sub Get_PVWave_Path()

        If (Not PVWave_RegistryII()) Then
            If (Not PVWave_Registry()) Then Exit Sub
        End If
        Dim path As String = System.Environment.GetEnvironmentVariable("path")
        Dim item As String
        Dim newpath = ""
        For Each item In path.Split(";")
            If Not item.ToLower.Contains("vni") Then newpath = String.Format("{0};{1}", newpath, item)
        Next
        newpath = String.Format("{0};{1}", newpath, Me.PVWPath)
        PVWaveLicense.LicensePath = System.IO.Path.Combine(PVWaveRootPath, "license")
        System.Environment.SetEnvironmentVariable("path", newpath)

    End Sub

    ' dgp rev 2/22/07
    ' dgp rev 3/1/07 check if the process is still running
    Public Sub Scan_Procs()

        Dim pees
        Dim AllProcs As Process

        Shell_List.Clear()
        Wave_List.Clear()

        pees = System.Diagnostics.Process.GetProcesses

        For Each AllProcs In pees
            If (AllProcs.ProcessName = cmd_name) Then Shell_List.Add(AllProcs)
            If (AllProcs.ProcessName = "wave") Then Wave_List.Add(AllProcs)
        Next

    End Sub

    ' dgp rev 4/13/07 Kill all Shell Processes
    Public Sub Kill_All()

        Scan_Procs()

        Dim Abort_Flag As Boolean = False
        Dim proc As Process
        ' are any shell processes running
        If (Wave_List.Count > 0) Then
            If (MsgBox("Flow Control is already running, abort pervious session?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
                For Each proc In Shell_List
                    If (Not proc.HasExited) Then proc.Kill()
                Next
                For Each proc In Wave_List
                    If (Not proc.HasExited) Then proc.Kill()
                Next
            End If
        End If

    End Sub

    ' dgp rev 4/12/07 Prepare the startup environment
    Public Function Prep_Startup() As Boolean

        Prep_Startup = False

        If (System.IO.Directory.Exists(FlowStructure.CurDistFullSpec)) Then

            Scan_Procs() ' create shell_list and wave_list

            Dim Abort_Flag As Boolean = False
            Dim proc As Process
            ' are any shell processes running
            If (Wave_List.Count > 0) Then
                If (MsgBox("Flow Control is already running, abort pervious session?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
                    For Each proc In Shell_List
                        If (Not proc.HasExited) Then proc.Kill()
                    Next
                    For Each proc In Wave_List
                        If (Not proc.HasExited) Then proc.Kill()
                    Next
                Else
                    Exit Function
                End If
            Else
                For Each proc In Shell_List
                    If (Not proc.HasExited) Then proc.Kill()
                Next
            End If

            If (Not System.IO.File.Exists(common_cmd)) Then Exit Function

            cmd_path = System.IO.Path.Combine(FlowStructure.CurDistFullSpec, cmd_name + ".exe")
            If (Not System.IO.File.Exists(cmd_path)) Then
                System.IO.File.Copy(common_cmd, cmd_path)
            End If

            Prep_Startup = True
        Else
            Prep_Startup = False
        End If


    End Function

    Private mWaveProc As Process
    ' dgp rev 2/14/2011
    Public ReadOnly Property WaveProc() As Process
        Get
            Return mWaveProc
        End Get
    End Property

    ' dgp rev 2/22/07
    Public Sub Get_Processes()

        proc_list = New Collection
        hndl_list = New Collection

        Dim pees
        Dim AllProcs As Process

        pees = System.Diagnostics.Process.GetProcesses

        For Each AllProcs In pees

            proc_list.Add(AllProcs.ProcessName, AllProcs.Id)
            hndl_list.Add(AllProcs.Id)
            If (AllProcs.ProcessName.ToLower.Contains("wave")) Then
                Dim name = AllProcs.ProcessName
                mWaveProc = AllProcs
                '                AddHandler mWaveProc.Exited, AddressOf WaveExited
            End If

        Next

    End Sub

    ' dgp rev 3/1/07 check if the process is still running
    Public Function Check_Processes(ByVal id As String) As Boolean

        Dim pees
        Dim AllProcs As Process

        Dim wave_flg As Boolean = False

        Check_Processes = False
        Shell_List.Clear()

        pees = System.Diagnostics.Process.GetProcesses

        For Each AllProcs In pees
            If (AllProcs.ProcessName = cmd_name) Then Shell_List.Add(AllProcs)
            If (AllProcs.ProcessName = "wave") Then wave_flg = True
        Next

        Dim proc As Process

        If (Not wave_flg) Then
            For Each proc In Shell_List
                proc.Kill()
            Next
        End If
        Check_Processes = True

    End Function

    ' dgp rev 5/14/09 process exited
    Public Sub ProcessExited(ByVal sender As Object, _
            ByVal e As System.EventArgs)

        RaiseEvent ExitedEvent(Now())

    End Sub

    ' dgp rev 5/14/09 process exited
    Public Sub WaveExited(ByVal sender As Object, _
            ByVal e As System.EventArgs)

        RaiseEvent ExitedEvent(Now())

    End Sub

    ' dgp rev 5/16/09 Parse the log file for errors
    Public Function Parse_Errors() As Boolean

        Dim Code_String As String
        Dim sr As StreamReader
        Dim file = Log_File
        Try
            sr = New StreamReader(Log_File)
        Catch ex As Exception
            Return False
        End Try

        Code_String = sr.ReadToEnd
        sr.Close()

        Dim objMatchCollection As MatchCollection
        ' dgp rev 3/26/08 let's first remove the comment fields
        '        Code_String = Regex.Replace(Code_String, ";.*\n", "")
        objMatchCollection = Regex.Matches(Code_String, "(\%.*error.*occur.*$)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        If objMatchCollection.Count > 0 Then
            Dim item = objMatchCollection.Count
        End If
        Return (objMatchCollection.Count > 0)

    End Function

    ' dgp rev 5/16/09 Parse the log file for errors
    Private Shared Function Parse_Errors(ByVal file As String) As Boolean

        Dim Code_String As String
        Dim sr As StreamReader
        Try
            sr = New StreamReader(file)
        Catch ex As Exception
            Return False
        End Try

        Code_String = sr.ReadToEnd
        sr.Close()

        Dim objMatchCollection As MatchCollection
        ' dgp rev 3/26/08 let's first remove the comment fields
        '        Code_String = Regex.Replace(Code_String, ";.*\n", "")
        objMatchCollection = Regex.Matches(Code_String, "(\%.*error.*occur.*$)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        If objMatchCollection.Count > 0 Then
            Dim item = objMatchCollection.Count
        End If
        Return (objMatchCollection.Count > 0)

    End Function

    ' dgp rev 5/16/09 
    Private Sub logchange(ByVal source As Object, ByVal e As  _
                        System.IO.FileSystemEventArgs)
        If (Parse_Errors(e.FullPath.ToString)) Then
            Dim err = "error"
        End If

    End Sub

    ' dgp rev 5/16/09 
    Private mWaveOn As Boolean = False

    ' dgp rev 5/16/09 
    Public Sub HandleCreationEvent(ByVal sender As Object, ByVal e As EventArrivedEventArgs)

        Dim ev As ManagementBaseObject = e.NewEvent
        If (CType(ev("TargetInstance"), ManagementBaseObject)("Name").ToString.ToLower.Contains("wave")) Then
            mWaveOn = True
            Dim a = ev("TargetInstance")
        End If
        Threading.Thread.Sleep(500)
        If (CType(ev("TargetInstance"), ManagementBaseObject)("Name").ToString.ToLower.Contains("wave")) Then
            mWaveOn = True
            Dim a = ev("TargetInstance")
        End If

    End Sub

    ' dgp rev 5/16/09 Process Deletion Handler, check for "wave.exe"
    Public Sub HandleDeletionEvent(ByVal sender As Object, ByVal e As EventArrivedEventArgs)

        Dim ev As ManagementBaseObject = e.NewEvent
        If (CType(ev("TargetInstance"), ManagementBaseObject)("Name").ToString.ToLower.Contains("wave")) Then
            mWaveOn = False
            RaiseEvent ExitedEvent("Wave.exe stopped")
        End If

    End Sub

    Private watcherDeletion As ManagementEventWatcher
    Private watcherCreation As ManagementEventWatcher

    ' dgp rev 2/14/2011
    Public Sub StopListening()

        Try
            ' Stop listening
            watcherCreation.Stop()
            watcherDeletion.Stop()
        Catch ex As Exception

        End Try

    End Sub

    ' dgp rev 2/14/2011
    Private Sub ProcessWatcher()

        'Watcher for Creation
        Dim queryCreation As WqlEventQuery = _
           New WqlEventQuery("__InstanceCreationEvent", _
           New TimeSpan(0, 0, 1), "TargetInstance isa ""Win32_Process""")
        watcherCreation = New  _
           ManagementEventWatcher(queryCreation)
        AddHandler watcherCreation.EventArrived, AddressOf HandleCreationEvent

        'Watcher for Deletion
        Dim queryDeletion As WqlEventQuery = New  _
           WqlEventQuery("__InstanceDeletionEvent", _
           New TimeSpan(0, 0, 1), "TargetInstance isa ""Win32_Process""")
        watcherDeletion = New  _
           ManagementEventWatcher(queryDeletion)
        AddHandler watcherDeletion.EventArrived, AddressOf HandleDeletionEvent

        ' Start listening
        watcherCreation.Start()
        watcherDeletion.Start()

    End Sub

    Private watchfolder As System.IO.FileSystemWatcher
    ' dgp rev 2/14/2011
    Private Sub StartWatcher()

        watchfolder = New System.IO.FileSystemWatcher()

        'this is the path we want to monitor
        watchfolder.Path = System.IO.Path.GetDirectoryName(Log_File.ToString)

        'Add a list of Filter we want to specify
        'make sure you use OR for each Filter as we need to
        'all of those 

        watchfolder.NotifyFilter = IO.NotifyFilters.DirectoryName
        watchfolder.NotifyFilter = watchfolder.NotifyFilter Or _
                                   IO.NotifyFilters.FileName
        watchfolder.NotifyFilter = watchfolder.NotifyFilter Or _
                                   IO.NotifyFilters.Attributes

        ' add the handler to each event
        AddHandler watchfolder.Changed, AddressOf logchange
        AddHandler watchfolder.Created, AddressOf logchange

        'Set this property to true to start watching
        watchfolder.EnableRaisingEvents = True

    End Sub

    ' dgp rev 9/20/2011 Launch routine to send out email of startup info
    Private Sub StartupEmailThread()

        Dim mReport As New EmailReporting
        mReport.SendVersionLog(FlowStructure.CurDistFullSpec)

    End Sub

    ' dgp rev 9/20/2011 Launch routine to send out email of startup info
    Private Sub LaunchStartupEmail()

        Dim objThread = New Thread(New ThreadStart(AddressOf StartupEmailThread))
        objThread.Name = "Log Version"
        objThread.Start()

    End Sub

    ' dgp rev 2/26/08
    Private m_Local_Ver = Nothing
    Public Property Local_Ver() As String
        Get
            If m_Local_Ver Is Nothing Then InitLocalVer()
            Return m_Local_Ver
        End Get

        ' dgp rev 2/26/08 only change setting if value is valid
        Set(ByVal value As String)
            If (Not Validate(value)) Then Return
            mLocalVerFlag = True
            If (mLocalVerFlag) Then
                mKey = "LocalVer"
                m_Local_Ver = value
                pvwXML.PutSetting(mKey, m_Local_Ver)
            End If
        End Set
    End Property


    ' dgp rev 7/17/07 Start the PVW process
    Public Sub Start_PVW()

        Create_Configs()

        If (Prep_Startup()) Then

            PVWCmd = New Process


            ' a new process is created from running the PV-Wave program
            PVWCmd.StartInfo.UseShellExecute = False
            ' redirect IO
            ' PVWCmd.StartInfo.RedirectStandardOutput = True
            '       PVWCmd.StartInfo.RedirectStandardError = True
            '      PVWCmd.StartInfo.RedirectStandardInput = True
            ' don't even bring up the console window
            PVWCmd.StartInfo.CreateNoWindow = True
            ' executable command line info
            cmd_str = cmd_path

            PVWCmd.StartInfo.FileName = cmd_path

            PVWCmd.StartInfo.EnvironmentVariables("pvwave_setup_flag") = "skip"

            '        PVWCmd.StartInfo.RedirectStandardInput = True

            '        PVWCmd.StartInfo.WorkingDirectory = Dist_Str
            PVWCmd.StartInfo.WorkingDirectory = FlowStructure.CurDistFullSpec
            Load_PVW_Attr()
            Log_Info("Flow Control Version")
            Log_Info(FlowStructure.CurDistFullSpec.ToString)

            PVWCmd.StartInfo.Arguments = mParams + " > """ + Log_File.ToString + """ 2>&1"

            '            PVWCmd.EnableRaisingEvents = True
            '   
            ' Add an event handler.
            '
            '           AddHandler PVWCmd.Exited, AddressOf Me.ProcessExited

            AddHandler WorkWatcher.NewWorkEventHandler, AddressOf NewWorkFileFound

            PVWCmd.Start()

            SystemLogAppend("Process started")
            '            AddHandler WorkWatcher.NewWorkEventHandler, AddressOf NewWorkFileFound
            WorkWatcher.StartWatching(FlowStructure.CurWork)
            '            PVWCmd.WaitForExit()
            LaunchStartupEmail()

            proc_name = PVWCmd.ProcessName
            ' dgp rev 7/18/07 reflect proc info later
            '            lblproc.Text = proc_name
            proc_id = PVWCmd.Id

            System.Threading.Thread.Sleep(1000)

            Get_Processes()

        End If

    End Sub

    ' dgp rev 8/31/06 create an XML settings file
    Public Function Write_XML(ByVal spec As String)

        Write_XML = False

        If (Not System.IO.Directory.Exists(spec)) Then Return Write_XML

        Dim NewRoot, NewChild
        Dim xdoc As New XmlDocument
        Dim NN As Xml.XmlNode
        Dim NewElement As XmlNode
        Dim NewInst As XmlProcessingInstruction

        NewInst = xdoc.CreateProcessingInstruction("xml", "version='1.0'")
        NewInst = xdoc.AppendChild(NewInst)

        NN = xdoc.CreateNode(XmlNodeType.DocumentType, "Environment", "")
        xdoc.AppendChild(NN)

        NewRoot = xdoc.CreateNode(XmlNodeType.Element, "Environment", "")
        NewRoot = xdoc.AppendChild(NewRoot)

        NewChild = xdoc.CreateNode(XmlNodeType.Element, "FlowControl", "")
        NewChild = NewRoot.appendChild(NewChild)
        NewChild = xdoc.CreateNode(XmlNodeType.Element, "Settings", "")
        NewChild = NewRoot.firstChild.appendChild(NewChild)

        'get the element interface to add attributes
        NewElement = NewChild

        ' dgp rev 11/1/06 only add work_path and setup_path to xml file
        Dim ie As IDictionaryEnumerator = PVW_Attr.GetEnumerator()
        Dim Attr As XmlAttribute
        While ie.MoveNext()
            Attr = xdoc.CreateAttribute(ie.Key)
            Attr.Value = ie.Value
            NewElement.Attributes.Append(Attr)
        End While

        Try
            xdoc.Save(spec)
            Write_XML = True
        Catch ex As Exception
            Write_XML = False
        End Try

    End Function

    ' dgp rev 8/31/06 create an XML settings file
    Public Function Write_XML()

        Write_XML = False

        Dim path As String = System.IO.Path.GetDirectoryName(FlowStructure.PVW_XML_File)
        If (Not System.IO.Directory.Exists(path)) Then Return Write_XML

        Dim NewRoot, NewChild
        Dim xdoc As New XmlDocument
        Dim NN As Xml.XmlNode
        Dim NewElement As XmlNode
        Dim NewInst As XmlProcessingInstruction

        NewInst = xdoc.CreateProcessingInstruction("xml", "version='1.0'")
        NewInst = xdoc.AppendChild(NewInst)

        NN = xdoc.CreateNode(XmlNodeType.DocumentType, "Environment", "")
        xdoc.AppendChild(NN)

        NewRoot = xdoc.CreateNode(XmlNodeType.Element, "Environment", "")
        NewRoot = xdoc.AppendChild(NewRoot)

        NewChild = xdoc.CreateNode(XmlNodeType.Element, "FlowControl", "")
        NewChild = NewRoot.appendChild(NewChild)
        NewChild = xdoc.CreateNode(XmlNodeType.Element, "Settings", "")
        NewChild = NewRoot.firstChild.appendChild(NewChild)

        'get the element interface to add attributes
        NewElement = NewChild

        ' dgp rev 11/1/06 only add work_path and setup_path to xml file
        Dim ie As IDictionaryEnumerator = PVW_Attr.GetEnumerator()
        Dim Attr As XmlAttribute
        While ie.MoveNext()
            Attr = xdoc.CreateAttribute(ie.Key)
            Attr.Value = ie.Value
            NewElement.Attributes.Append(Attr)
        End While

        Try
            xdoc.Save(FlowStructure.PVW_XML_File)
            Write_XML = True
        Catch ex As Exception
            Write_XML = False
        End Try

    End Function

    ' dgp rev 2/20/08 initialize instance
    Private Sub Init()

        ' dgp rev 7/17/07 move the creation of user path into class
        Utility.Create_Tree(Utility.MyAppPath)

        Username = System.Environment.GetEnvironmentVariable("username")

        Dim uname As String = "PVW_" + Format(Now(), "yyMMddhhmm")

        Log_File = System.IO.Path.Combine(Utility.MyAppPath, uname + ".log")

        Get_Dynamic()


        ' dgp rev 3/30/2011 don't load the attributes, wait until they are needed
        '        Load_PVW_Attr()
        ' dgp rev 6/25/09 if no access to registry, wave application may need explicit setting
        output_str = System.IO.Path.Combine(Utility.MyAppPath, "wave.log")

        ProcessWatcher()

    End Sub

    ' dgp rev 2/22/07 create a new instance - scan for root
    Public Sub New()

        Init()

    End Sub

    ' dgp rev 2/22/07 create a new instance - assign root
    Public Sub New(ByVal FullPath As String)

        FlowStructure.Set_Root(FullPath)
        Init()

    End Sub

    ' dgp rev 2/14/2011
    Protected Overrides Sub Finalize()

        ' Start listening
        Try
            '            watcherCreation.Stop()
        Catch ex As Exception

        End Try
        Try
            '            watcherDeletion.Stop()
        Catch ex As Exception

        End Try

        MyBase.Finalize()


    End Sub

    ' dgp rev 2/14/2011 testing
    Public Sub TestClusterXML()

        ClusterStats.ClusterStatFile = "F:\FlowRoot\Users\plugged\Work\20101215101931\20110209042730\stat_table.xml"
        If ClusterStats.ProcessClusters() Then

        End If

    End Sub

    Public Delegate Sub WorkEventHandler(ByVal SomeString As String)
    Public Shared Event WorkEvent As WorkEventHandler

    Private Sub NewWorkFileFound(FileName As String)

        Console.WriteLine(FileName)

    End Sub

End Class
