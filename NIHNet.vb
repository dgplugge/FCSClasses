' Author: Donald G. Plugge
' Date: 2/12/09
' Class: NIH Network specific information
Imports System.Management
'Imports System.Web.Security
Imports System.Security.Principal

Public Class NIHNet

    Public Shared wmiOptions As Management.ConnectionOptions
    Public Shared strComputerName As String = "NT-EIB-10-6B16"

    Private Shared mSharePath As ArrayList
    Private Shared mShareName As ArrayList

    Public Delegate Sub FSRemoteEventHandler(ByVal SomeString As String)
    Public Shared Event FSRemoteEvent As FSRemoteEventHandler

    ' dgp rev 1/18/07
    Private Function Get_DiskSpace() As Collection

        Dim dyn_path As String
        dyn_path = "\\" + FlowServer.FlowServer + "\root\cimv2"

        Dim mClass As ManagementClass = New ManagementClass("Win32_LogicalDisk")
        Dim mPath2 As New ManagementPath(dyn_path)
        Dim myScope As New ManagementScope(mPath2)

        Dim Parts As New Collection
        '* connect to WMI namespace
        Try
            myScope.Connect()
        Catch ex As Exception
            Return Parts
        End Try

        If myScope.IsConnected = False Then
            Parts.Add("No device")
        Else
            'Dim QStr2 As String = "Select PartComponent from Win32_GroupUser"

            Dim wmi As System.Management.ManagementClass
            Dim obj As System.Management.ManagementObject
            Dim ans As Double

            ' The argument to ManagementClass includes 3 things:
            ' 1) the name of remote PC
            ' 2) the WMI "Namespace" (\root\cimv2:)
            ' 3) the WMI Class (Win32_LogicalDisk)
            wmi = New System.Management.ManagementClass("\\" & FlowServer.FlowServer & "\root\cimv2:" & _
             "Win32_LogicalDisk")

            For Each obj In wmi.GetInstances()
                If obj("DeviceID").ToString = "C:" Then
                    ans = Convert.ToDouble(obj("FreeSpace")) / (1024.0# * 1024.0#)
                    Exit For
                End If
            Next


            Dim QStr2 As String = "SELECT *  FROM Win32_LogicalDisk WHERE  DriveType = ""3"""
            QStr2 = QStr2 + "Win32_Group.Domain=""10-plugge-1"
            QStr2 = QStr2 + """,Name=""Flow Users"""

            '            dim QStr2 as String = "SELECT partcomponent FROM Win32_GroupUser WHERE GroupComponent=\"Win32_Group.Domain=\"10-plugge-1\",Name=\"Flow Users\""
            Dim mQuery As New ObjectQuery(QStr2)
            Dim mOpt As New ObjectGetOptions
            Dim mSearcher As New ManagementObjectSearcher(myScope, mQuery)

            mClass.Scope = myScope

            Dim mObj As ManagementObject
            Dim PC() As String

            Parts.Add("Select User")
            For Each mObj In mSearcher.Get
                PC = mObj("PartComponent").ToString.Split("=")
                Parts.Add(PC(PC.Length - 1).Trim(""""))
            Next

            Dim oMan As New ManagementObject

        End If
        Return Parts

    End Function


    Public Shared Function GetSpace(ByVal drv As String) As Integer

        Dim ms As System.Management.ManagementScope
        Dim oq As System.Management.ObjectQuery
        Dim mos As System.Management.ManagementObjectSearcher
        Dim obj As System.Management.ManagementObject
        Dim ans As Double

        ' The "scope" includes the name of the PC and the WMI namespace
        ms = New System.Management.ManagementScope("\\" & FlowServer.FlowServer.ToString & _
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

    Private Shared mCurShare As String = "Upload"
    Public Shared Property CurShare As String
        Get
            Return mCurShare
        End Get
        Set(ByVal value As String)
            mCurShare = value
        End Set
    End Property

    ' dgp rev 6/1/2011 Get Mount Space 
    Public Shared Function CheckMountPoint(ByVal mnt As String) As Boolean

        Dim ms As System.Management.ManagementScope
        Dim oq As System.Management.ObjectQuery
        Dim mos As System.Management.ManagementObjectSearcher
        Dim obj As System.Management.ManagementObject

        ' The "scope" includes the name of the PC and the WMI namespace
        ms = New System.Management.ManagementScope("\\" & FlowServer.FlowServer.ToString & _
         "\root\cimv2")

        CheckMountPoint = False
        ' use WQL to get just the one instance we want.  This should look familiar
        ' to those who are used to SQL
        oq = New System.Management.ObjectQuery("select * from Win32_Volume where Label = '" + mnt + "'")
        ' execute the query
        mos = New System.Management.ManagementObjectSearcher(ms, oq)
        For Each obj In mos.Get
            mFreeSpace = Convert.ToDouble(obj("FreeSpace"))
            mTotalSpace = Convert.ToDouble(obj("Capacity"))
            mFreeRatio = mFreeSpace / mTotalSpace
            mValidSpace = True
            CheckMountPoint = True
            Exit For
        Next

    End Function

    ' dgp rev 9/13/2010
    Public Shared Function ScanShares() As Boolean

        If FCS_Classes.NIHNet.ShareExists(mCurShare) Then

            If Not FCS_Classes.NIHNet.CheckMountPoint(mCurShare) Then
                Dim dev = FCS_Classes.NIHNet.GetShare(mCurShare)
                mFreeRatio = FCS_Classes.NIHNet.GetFreeRatio(dev)
            End If
        Else
            mFreeRatio = -1
            Return False
        End If
        RaiseEvent FSRemoteEvent("xyzzy")

        Return True

    End Function

    Private Shared mFreeSpace As Double
    Private Shared mTotalSpace As Double
    Private Shared mCurDevice As String
    Private Shared mValidSpace As Boolean = False

    ' dgp rev 5/13/2011 
    Public Shared Sub CalcDiskSpace(ByVal device As String)

        mValidSpace = False
        mCurDevice = device

        mFreeRatio = 0.0

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "select * from Win32_LogicalDisk where DriveType=3 and DeviceId = '" + device.ToString.Substring(0, 2) + "'"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.FlowServer.ToString & _
         "\root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            mFreeSpace = management_object("FreeSpace")
            mTotalSpace = management_object("Size")
            mFreeRatio = mFreeSpace / mTotalSpace
            mValidSpace = True
        Next management_object

    End Sub

    ' dgp rev 5/13/2011
    Private Shared mFreeRatio
    Public Shared ReadOnly Property FreeRatio As Double
        Get
            If mValidSpace Then Return mFreeRatio
            Return -1.0
        End Get
    End Property

    ' dgp rev 5/13/2011
    Public Shared Function GetTotalSpace() As Double

        If mValidSpace Then Return mTotalSpace
        Return -1.0

    End Function


    ' dgp rev 5/13/2011
    Public Shared Function GetTotalSpace(ByVal device) As Double

        CalcDiskSpace(device)
        Return GetTotalSpace()

    End Function

    ' dgp rev 5/13/2011
    Public Shared Function GetFreeSpace() As Double

        If mValidSpace Then Return mFreeSpace
        Return -1.0

    End Function


    ' dgp rev 5/13/2011
    Public Shared Function GetFreeSpace(ByVal device) As Double

        CalcDiskSpace(device)
        Return GetFreeSpace()

    End Function

    ' dgp rev 5/13/2011
    Public Shared Function GetFreeRatio() As Double

        If mValidSpace Then Return mFreeRatio
        Return -1.0

    End Function


    ' dgp rev 5/13/2011
    Public Shared Function GetFreeRatio(ByVal device) As Double

        CalcDiskSpace(device)
        Return GetFreeRatio()

    End Function

    ' dgp rev 9/13/2010
    Public Shared Function GetSharePath(ByVal name) As String

        mSharePath = New ArrayList
        mShareName = New ArrayList

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_Share"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.FlowServer.ToString & _
         "\root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            mShareName.Add(management_object("Name"))
            mSharePath.Add(management_object("Path"))
        Next management_object

        If mShareName.IndexOf(name) = -1 Then Return ""
        Return mSharePath(mShareName.IndexOf(name))

    End Function

    Private Shared mRemoteXMLPath As String

    Private Shared mPath As String
    Private Shared mServerName As String = "NT-EIB-10-6B16"
    Private Shared mScope As Management.ManagementScope
    Private Shared mClass As Management.ManagementClass

    Private Shared Function CreateShare(ByVal name, ByVal path) As Boolean

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

    ' dgp rev 5/17/2011 Exper XML Path
    Public Shared Function ExperXMLPath() As String

        ExperXMLPath = ""
        If ShareExists(FlowServer.ExperXMLShare) Then
            mRemoteXMLPath = FlowServer.ExperXMLShare
        Else
            If CreateShare(FlowServer.ExperXMLShare, "I:\George") Then
                mRemoteXMLPath = FlowServer.ExperXMLShare
            Else
                MsgBox("Error creating share", MsgBoxStyle.Information)
            End If
        End If

    End Function

    ' dgp rev 9/13/2010
    Public Shared Function ShareExists(ByVal name) As Boolean

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_Share"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.FlowServer.ToString & _
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

    ' dgp rev 9/13/2010
    Public Shared Function GetShare(ByVal name) As String

        mSharePath = New ArrayList
        mShareName = New ArrayList

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_Share"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.FlowServer.ToString & _
         "\root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            mShareName.Add(management_object("Name"))
            mSharePath.Add(management_object("Path"))
        Next management_object

        If mShareName.IndexOf(name) = -1 Then Return ""
        Return System.IO.Path.GetPathRoot(mSharePath(mShareName.IndexOf(name)))

    End Function

    ' dgp rev 9/13/2010
    Public Shared Function GetUploadDevice() As String

        mSharePath = New ArrayList
        mShareName = New ArrayList

        Dim query As ManagementObjectSearcher
        Dim queryCollection As ManagementObjectCollection

        Dim query_command As String = "SELECT * FROM Win32_Share"

        Dim msc As ManagementScope = New ManagementScope("\\" & FlowServer.FlowServer.ToString & _
         "\root\cimv2")

        Dim select_query As SelectQuery = New SelectQuery(query_command)

        query = New ManagementObjectSearcher(msc, select_query)
        queryCollection = query.Get()

        Dim management_object As ManagementObject

        For Each management_object In queryCollection
            mShareName.Add(management_object("Name"))
            mSharePath.Add(management_object("Path"))
        Next management_object

        Dim path = mSharePath(mShareName.IndexOf("Upload"))
        Return System.IO.Path.GetPathRoot(path)

    End Function


    ' dgp rev 1/18/07
    Public Shared Sub CheckOperator()

        Dim mCIMV2 As String = "\\" & strComputerName & "\root\cimv2:Win32_GroupUser"
        Dim mClass As ManagementClass = New ManagementClass("Win32_GroupUser")
        Dim myScope = New ManagementScope(mCIMV2, wmiOptions)

        mFlowLabUsers = New ArrayList

        '* connect to WMI namespace
        ' attempt connection
        Try
            myScope.Connect()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
            Exit Sub
            ' failure
        End Try

        Dim Parts As New Collection

        If myScope.IsConnected = False Then
            ' connection failed
        Else
            ' successful connection
            Dim QStr2 As String = "Select PartComponent from Win32_GroupUser where GroupComponent=""Win32_Group.Domain='NT-EIB-10-6B16',Name='Flow User'"""
            Dim mQuery As New ObjectQuery(QStr2)

            Dim mSearcher As New ManagementObjectSearcher(myScope, mQuery)

            mClass.Scope = myScope

            Dim PC() As String
            Dim mObj As ManagementObject
            ' loop thru the results and parse out the username without quotes
            For Each mObj In mSearcher.Get
                PC = mObj("PartComponent").ToString.Split("=")
                mFlowLabUsers.Add(PC(2).ToString.Replace("""", ""))
            Next
            If (mFlowLabUsers.Count > 1) Then mFlowLabUsers.Sort()

        End If

    End Sub


    ' dgp rev 1/18/07
    Public Shared Sub Get_NCIUsers()

        Dim mCIMV2 As String = "\\" & strComputerName & "\root\cimv2:Win32_GroupUser"
        Dim mClass As ManagementClass = New ManagementClass("Win32_GroupUser")
        Dim myScope = New ManagementScope(mCIMV2, wmiOptions)

        mFlowLabUsers = New ArrayList

        Try

            '* connect to WMI namespace
            ' attempt connection
            myScope.Connect()

            Dim Parts As New Collection

            If myScope.IsConnected = False Then
                ' connection failed
            Else
                ' successful connection
                Dim QStr2 As String = "Select PartComponent from Win32_GroupUser where GroupComponent=""Win32_Group.Domain='NT-EIB-10-6B16',Name='Flow User'"""
                Dim mQuery As New ObjectQuery(QStr2)

                Dim mSearcher As New ManagementObjectSearcher(myScope, mQuery)

                mClass.Scope = myScope

                Dim PC() As String
                Dim mObj As ManagementObject
                ' loop thru the results and parse out the username without quotes
                For Each mObj In mSearcher.Get
                    PC = mObj("PartComponent").ToString.Split("=")
                    mFlowLabUsers.Add(PC(2).ToString.Replace("""", ""))
                Next
                If (mFlowLabUsers.Count > 1) Then mFlowLabUsers.Sort()

            End If
        Catch ex As Exception
            mFlowLabUsers.Add(Environment.UserName)
        End Try

    End Sub

    ' dgp rev 9/28/2010 Flow Lab Users
    Private Shared mFlowLabUsers As ArrayList = Nothing
    Public Shared ReadOnly Property FlowLabUsers As ArrayList
        Get
            If mFlowLabUsers Is Nothing Then Get_NCIUsers()
            Return mFlowLabUsers
        End Get
    End Property


End Class
