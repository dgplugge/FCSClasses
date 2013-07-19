' Name:     Flow Root Class
' Author:   Donald G Plugge
' Date:     7/20/07
' Purpose:  Class to maintain the local Flow Root structure
Imports HelperClasses

Public Class FlowRoot

    Private mUsername As String = System.Environment.GetEnvironmentVariable("username")

    Private mSep As Char = System.IO.Path.DirectorySeparatorChar
    Private mLastLocal As String
    Private m_Run_List As Collection

    ' one time initialization of FlowRoot
    Private mDrive As String
    Private mDrives As ArrayList
    Private mRoot As String
    Private mRoots As ArrayList

    Private mServer As String = "NT-EIB-10-6B16"
    Private mShare_Flow As String = "root2"
    Private mRemoteRoot As String

    Private mXML As New Dynamic("FlowRoot")
    ' dgp rev 2/19/08 error messages 
    Private mMessage As String

    Public ReadOnly Property Message() As String
        Get
            Return mMessage
        End Get
    End Property

    Public ReadOnly Property RemoteRoot() As String
        Get
            Return String.Format("\\{0}\{1}\Users\{2}", mServer, mShare_Flow, mUsername)
        End Get
    End Property

    ' dgp rev 7/20/07 
    Private Sub ScanRoot()

        Dim tstDrv As String
        mRoots = New ArrayList
        mDrives = New ArrayList

        Dim drv
        ' dgp rev 2/19/08 scan drives for valid locations
        For Each drv In Utility.LocalDrives
            mDrives.Add(drv.DriveLetter)
            tstDrv = drv.DriveLetter + ":" + mSep + "FlowRoot"
            If System.IO.Directory.Exists(tstDrv) Then mRoots.Add(tstDrv)
        Next

    End Sub
    ' dgp rev 7/20/07 Scan for User Root
    Private Function ScanUser() As Boolean

        mUserRoot = ""
        If (RootList.Count = 0) Then Exit Function
        Dim item
        Dim tstPath As String
        For Each item In RootList
            tstPath = system.io.path.combine(system.io.path.combine(item, "Users"), mUsername)
            If System.IO.Directory.exists(tstPath) Then mUserRoot = tstPath
        Next

    End Function
    ' dgp rev 7/20/07 Scan for User Root
    Private mUserRoot As String
    Public ReadOnly Property UserRoot() As String
        Get
            If (mUserRoot Is Nothing) Then ScanUser()
            Return mUserRoot
        End Get
    End Property
    ' dgp rev 7/20/07
    Private mValid As Boolean = False
    Public ReadOnly Property Valid()
        Get
            Return (RootList.Count > 0)
        End Get
    End Property
    ' dgp rev 7/20/07 
    Public ReadOnly Property RootList() As ArrayList
        Get
            If (mRoots Is Nothing) Then ScanRoot()
            Return mRoots
        End Get
    End Property

    ' dgp rev 2/19/08 FlowRoot is read from XML file
    Private mFlowRoot As String
    Public ReadOnly Property FlowRoot() As String
        Get
            Return mFlowRoot
        End Get
    End Property

    ' dgp rev 5/7/08 Find Flow Root
    Public Function Create_SubRoot(ByVal SubDir As String) As Boolean

        Create_SubRoot = False

        Dim TstPath As String

        If (RootList.Count = 0) Then Exit Function

        Dim item
        For Each item In RootList
            TstPath = system.io.path.combine(item, "Subdir")
            If System.IO.Directory.exists(TstPath) Then Exit Function
        Next

        ' dgp rev 2/12/08 replace System.IO.Directory.CreateDirectory with Global Helper Create_Tree
        TstPath = system.io.path.combine(RootList.Item(RootList.Count - 1), SubDir)
        Create_SubRoot = Utility.Create_Tree(TstPath)
        If (Create_SubRoot) Then m_Data_Root = TstPath

    End Function

    ' dgp rev 5/9/07 Scan for particular subdir
    Public Sub Find_SubRoot(ByVal SubDir As String)

        Dim tstdrv As String
        If (RootList.Count = 0) Then Exit Sub

        Dim item
        For Each item In RootList
            tstdrv = item + mSep + SubDir
            If (System.IO.Directory.exists(tstdrv)) Then m_Data_Root = tstdrv
        Next

    End Sub

    Private mUsers As String

    ' dgp rev 5/9/07 Scan For Settings
    Public Sub Find_Settings()

        Dim item
        Dim tstPath As String

        For Each item In RootList
            tstPath = system.io.path.combine(item, "Users")
            If (System.IO.Directory.exists(tstPath)) Then
                mUsers = tstPath
                tstPath = system.io.path.combine(tstPath, mUsername)
                If (System.IO.Directory.exists(tstPath)) Then
                    mUserRoot = tstPath
                    tstPath = system.io.path.combine(tstPath, "Settings")
                    If (System.IO.Directory.exists(tstPath)) Then
                        mSettings = tstPath
                    End If
                End If
            End If
        Next

    End Sub
    ' dgp rev 7/20/07
    Private Sub Create_Settings()

        Dim tstPath As String = ""
        mSettings = ""

        If (mUsers Is Nothing) Then mUsers = system.io.path.combine(RootList.Item(RootList.Count - 1), "Users")
        If (mUserRoot Is Nothing) Then mUserRoot = system.io.path.combine(mUsers, mUsername)

        tstPath = system.io.path.combine(mUserRoot, "Settings")

        Utility.Create_Tree(tstPath)

        If (System.IO.Directory.exists(tstPath)) Then mSettings = tstPath

    End Sub
    ' dgp rev 5/17/07 Data Root
    Private mSettings As String
    Public Property Settings() As String
        Get
            If (mSettings Is Nothing) Then Find_Settings()
            If (mSettings Is Nothing) Then Create_Settings()
            Return mSettings
        End Get
        Set(ByVal value As String)
            If (mSettings = value) Then Exit Property
            mSettings = value
        End Set
    End Property


    ' dgp rev 5/17/07 Data Root
    ' dgp rev 11/28/07 get the DataRoot from one place
    ' dgp rev 2/19/08 how do we know if data is set, existant and valid
    Private m_Data_Root As String
    Private mDataRootFlag As Boolean = False

    Public Property Data_Root() As String
        Get
            If (mDataRootFlag) Then Find_SubRoot("Data")
            If (mDataRootFlag) Then Create_SubRoot("Data")
            Return m_Data_Root
        End Get
        Set(ByVal value As String)
            If (m_Data_Root = value) Then Exit Property
            m_Data_Root = value
        End Set
    End Property

    ' dgp rev 2/20/08 Does Root Exist
    Private Function Check_Root() As Boolean

        Check_Root = False
        If (mRoots.Count > 0) Then
            Check_Root = True
            mFlowRoot = mRoots.Item(mRoots.Count - 1)
        End If

    End Function

    ' dgp rev 2/20/08 Populate a new Root with the proper subfolders
    Private Function Populate_Root() As Boolean

        Populate_Root = False

        Dim path As String

        path = mFlowRoot

        If (Not Utility.Create_Tree(system.io.path.combine(path, "Data"))) Then MsgBox("Failed to create Data area", MsgBoxStyle.Information)
        If (Not Utility.Create_Tree(system.io.path.combine(path, "Distribution"))) Then MsgBox("Failed to create Distribution area", MsgBoxStyle.Information)

        path = Environment.GetEnvironmentVariable("Username")
        path = system.io.path.combine(system.io.path.combine(mFlowRoot, "Users"), path)

        If (Not Utility.Create_Tree(system.io.path.combine(path, "Work"))) Then MsgBox("Failed to create Work area", MsgBoxStyle.Information)
        If (Not Utility.Create_Tree(system.io.path.combine(path, "Settings"))) Then
            MsgBox("Failed to create Settings area", MsgBoxStyle.Information)
        Else
            Populate_Root = True
        End If


    End Function

    ' dgp rev 2/20/08 Create the FlowRoot
    Private Function Create_Root() As Boolean

        Create_Root = False
        ' dgp rev 2/19/08 review valid locations and select or create FlowRoot
        If (mDrives.Count > 0) Then
            Dim tst As String
            tst = mDrives.Item(mDrives.Count - 1) + ":" + mSep + "FlowRoot"
            If (Utility.Create_Tree(tst)) Then
                mFlowRoot = tst
                If (Populate_Root()) Then
                    Create_Root = True
                End If
            Else
                mMessage = "Root Creation Failed - " + tst
            End If
        Else
            mMessage = "No valid local drives"
        End If

    End Function

    ' dgp rev 2/19/08 First time running FlowRoot, must initialize
    Private Function Init_Root() As Boolean

        Init_Root = False

        ' dgp rev 7/20/07 
        ScanRoot()
        If (Check_Root()) Then
            Init_Root = True
        Else
            If (Create_Root()) Then
                Init_Root = True
            End If
        End If

    End Function

    ' dgp rev 2/20/08 Establish FlowRoot
    ' Sets FlowRoot variable and saves in XML file.  
    Private Function Establish_Root() As Boolean

        Establish_Root = False
        ' dgp rev 2/20/08 keep this simple, if XML settings exists, then read it.
        If (mXML.NewFlag) Then
            mMessage = "One Time Initialization"
            ' one-time initialization
        Else
            ' read previous setup and validate
            If (mXML.Exists("FlowRoot")) Then
                If (System.IO.Directory.exists(mXML.GetSetting("FlowRoot"))) Then
                    Establish_Root = True
                    mFlowRoot = mXML.GetSetting("FlowRoot")
                    Exit Function
                Else
                    MsgBox("FlowRoot no longer valid", MsgBoxStyle.Exclamation)
                End If
            End If
        End If

        If (Init_Root()) Then
            Establish_Root = True
            mXML.PutSetting("FlowRoot", mFlowRoot)
            ScanRoot()
            MsgBox("Successful One-Time Initialization of PC", MsgBoxStyle.Information)
        Else
            MsgBox("Initialization Failed - " + mMessage, MsgBoxStyle.Exclamation)
        End If

    End Function

    ' dgp rev 7/20/07
    Public Sub New(ByVal path As String)

        If (System.IO.Directory.exists(path)) Then
            mFlowRoot = path
        Else
            MsgBox("Failed to establish specific FlowRoot - " + path, MsgBoxStyle.Information)
        End If

    End Sub

    ' dgp rev 7/20/07
    Public Sub New()

        If (Not Establish_Root()) Then
            MsgBox("Failed to establish FlowRoot", MsgBoxStyle.Information)
        End If

    End Sub

End Class
