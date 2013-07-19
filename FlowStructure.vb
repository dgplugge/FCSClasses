' Name:     Flow Structure Module   
' Author:   Donald G Plugge
' Date:     2/22/08
' Purpose:  Module to facilitate the local Flow Root structure
' 
Imports HelperClasses
Imports System.IO
Imports System.Xml

' dgp rev 2/26/08 what if no FlowRoot?  Create it, along with Distribution,
' Work, Data, and Settings
' dgp rev 2/26/08 what if no data?  Flow Control must be able to run without data

' dgp rev 7/15/08 Structure for each component of FlowRoot
' 1) mXXXXScan     does a scan need to be performed
' 2) mXXXXAny      does anything exist in the root - if empty then newly created
' 3) mXXXXRoots    a list of all roots - valid roots must contain info
' 4) mCurXXXX      current setting - what is the persistent valid?
' 5) mCurXXXXValid current valid - is the current setting still valid?
' 6) CreateXXXX    create the empty structure 
' 7) PopulateXXXX  populate with defaults (perhaps part of create)

Public Class FlowStructure

    ' use delegates for Repeat

    Public Delegate Sub FSRemoteEventHandler(ByVal SomeString As String)
    Public Shared Event FSRemoteEvent As FSRemoteEventHandler

    Public Delegate Sub TransferEventHandler(ByVal Count As Int16)
    Public Shared Event TransferEvent As TransferEventHandler

    Public Delegate Sub TransferSourceHandler(ByVal SomeString As String)
    Public Shared Event TransferSource As TransferSourceHandler

    Public Delegate Sub TransferTargetHandler(ByVal SomeString As String)
    Public Shared Event TransferTarget As TransferTargetHandler

    Private Shared mLogger As Logger
    Public Shared Property Logger() As Logger
        Get
            Return mLogger
        End Get
        Set(ByVal value As Logger)
            mLogger = value
        End Set
    End Property

    Public Shared Sub Log_Info(ByVal text As String)

        If (mLogger Is Nothing) Then Exit Sub

        mLogger.Log_Info(text)

    End Sub

    ' dgp rev 2/19/08 FlowRoot is read from XML file
    Private Shared mFlowRoot As String

    ' dgp rev 5/17/07 Data Root
    ' dgp rev 11/28/07 get the DataRoot from one place
    ' dgp rev 2/19/08 how do we know if data is set, existant and valid
    Private Shared mUserRoot As String
    Private Shared mUserFlag As Boolean = False

    Private Shared mSettingRoot As String
    Private Shared mSettingRoots As ArrayList
    Private Shared mSettingScan As Boolean = False
    Private Shared mSettingExists As Boolean = False

    ' dgp rev 2/22/08 remote server configuration information
    Private Shared mUsername As String = System.Environment.GetEnvironmentVariable("username")
    Private Shared mDomain As String = System.Environment.GetEnvironmentVariable("userdomain")
    Private Shared mServer As String = "NT-EIB-10-6B16"
    Private Shared mShare_Flow As String = "root2"
    Private Shared mServerValid As Boolean = False

    Private Enum mFolderFlag
        [None] = 0
        Empty = 1
        Folders = 2
        Files = 4
    End Enum

    Private Enum mState
        [None] = 0
        Empty = 1
        Populated = 2
    End Enum

    Private Shared mNoRoot As Boolean = True
    Private Shared mNoData As Boolean = True
    Private Shared mNoDepot As Boolean = True
    Private Shared mNoUser As Boolean = True
    Private Shared mNoWork As Boolean = True
    Private Shared mNoSettings As Boolean = True

    ' dgp rev 7/10/08 Mask of installation problems
    Private Enum mFlowAnalysis
        [Valid] = 0
        NoRoot = 1
        NoData = 2
        NoDist = 4
        NoUsers = 8
        NoWork = 16
        NoSettings = 32
    End Enum

    Private Shared mCurState As mFlowAnalysis = mFlowAnalysis.Valid

    Public Shared ReadOnly Property NoRoot() As Boolean
        Get
            Return mNoRoot
        End Get
    End Property

    Public Shared ReadOnly Property NoData() As Boolean
        Get
            Return mNoData
        End Get
    End Property

    Public Shared ReadOnly Property NoWork() As Boolean
        Get
            Return mNoWork
        End Get
    End Property

    Public Shared ReadOnly Property NoDist() As Boolean
        Get
            If System.IO.Directory.Exists(FlowStructure.FlowRoot) Then
                Dim path = System.IO.Path.Combine(FlowStructure.FlowRoot, "Distribution")
                If System.IO.Directory.Exists(path) Then
                    If CurDist IsNot Nothing Then
                        path = System.IO.Path.Combine(path, CurDist)
                        Return Not System.IO.Directory.Exists(path)
                    End If
                End If
            End If
            Return True

        End Get
    End Property

    Public Shared ReadOnly Property NoSettings() As Boolean
        Get
            Return mNoSettings
        End Get
    End Property

    Public Shared ReadOnly Property CurState() As Int16
        Get
            Return (Not (mNoRoot Or NoDist Or mNoData Or mNoWork Or mNoSettings))
        End Get
    End Property

    ' dgp rev 7/14/08 Set Mask
    Private Shared Sub Init_Mask()

        mNoRoot = True
        mNoData = True
        mNoWork = True
        mNoUser = True
        mNoSettings = True

    End Sub



    ' dgp rev 12/1/09 
    Public Shared Function Validate() As Boolean

        Bit_Onoff(mFlowAnalysis.NoRoot, (FlowRoot IsNot Nothing))
        Bit_Onoff(mFlowAnalysis.NoData, (Data_Root IsNot Nothing))
        Bit_Onoff(mFlowAnalysis.NoDist, (Dist_Root IsNot Nothing))
        Bit_Onoff(mFlowAnalysis.NoSettings, (Settings IsNot Nothing))
        Bit_Onoff(mFlowAnalysis.NoUsers, (User_Root IsNot Nothing))
        Bit_Onoff(mFlowAnalysis.NoWork, (Work_Root IsNot Nothing))

        Return mFlowAnalysis.Valid = 0

    End Function

    ' dgp rev 7/14/08 Set Mask
    Private Shared Sub Bit_Onoff(ByVal bit As mFlowAnalysis, ByVal val As Boolean)

        Select Case bit
            Case mFlowAnalysis.NoRoot
                mNoRoot = val
            Case mFlowAnalysis.NoData
                mNoData = val
            Case mFlowAnalysis.NoDist
            Case mFlowAnalysis.NoWork
                mNoWork = val
            Case mFlowAnalysis.NoUsers
                mNoUser = val
            Case mFlowAnalysis.NoSettings
                mNoSettings = val
        End Select

    End Sub

    ' dgp rev 7/10/08 flags for existance of each item in FlowRoot Structure
    Private Shared mSettingFlag As mFolderFlag
    Private Shared mUsersFlag As mFolderFlag

    ' dgp rev 2/22/08 The program distribution root
    ' dgp rev 2/26/08 keep must of the distribution work inside this class
    ' dgp rev 2/27/08 
    ' 1) mXXXXScan     does a scan need to be performed
    ' 2) mXXXXState    state of current none, empty, populated 
    ' 3) mXXXXRoot     a single current root - valid root must contain info
    ' 4) mXXXXRoots    a list of all roots - valid roots must contain info
    ' 5) mCurXXXX      current setting - what is the persistent valid?
    ' 6) mCurXXXXValid current valid - is the current setting still valid?
    ' 7) mLatestXXXX   latest to be created
    ' 8) CreateXXXX    subroutine - create the empty structure 
    ' 9) PopulateXXXX  subroutine - populate with defaults (perhaps part of create)
    Private Shared mDistScan As Boolean = False
    Private Shared mDistState As mState
    Private Shared mDistRoot As String
    Private Shared mCurDistValid As Boolean

    Private Shared mWorkRoots As ArrayList
    Private Shared mWorkScan As Boolean = False
    Private Shared mWorkExists As Boolean = False
    Private Shared mWorkState As mState
    Private Shared mCurWork = Nothing
    Private Shared mCurWorkValid As Boolean
    Private Shared mLatestWork As String

    Private Shared mSettingsRoot As String
    Private Shared mSettingsRoots As ArrayList
    Private Shared mSettingsScan As Boolean = False
    Private Shared mSettingsState As mState
    Private Shared mCurSettings As String
    Private Shared mCurSettingsValid As Boolean
    Private Shared mLatestSettings As String

    ' dgp rev 5/21/08 create a run array with data sorted according to criteria
    Private Shared mRunArray As ArrayList
    ' dgp rev 4/28/09 current display order for runs
    Private Shared mRunDisplay As ArrayList
    Private Shared mRunScanFlag As Boolean = True
    Private Shared mRunCount As Boolean = False

    Private Shared mCurRun = Nothing
    ' dgp rev 7/6/09 Remove Current Run
    Public Shared Function RemoveBrowseRun() As Boolean

        Dim path = BrowseRun.Data_Path
        If System.IO.Directory.Exists(path) Then
            mCurRun = Nothing
            Try
                Utility.DeleteTree(path)
            Catch ex As Exception

            End Try
        End If

    End Function

    ' dgp rev 4/28/09 Run Display - sorted or filtered run array
    Public Shared ReadOnly Property RunDisplay() As ArrayList
        Get
            If mRunDisplay Is Nothing Then mRunDisplay = RunArray
            Return mRunDisplay
        End Get
    End Property
    ' dgp rev 7/1/09 
    Public Shared Function Server_Data() As Boolean

        Server_Data = False
        Dim Server = String.Format("\\{0}", mServer)

        Dim path = System.IO.Path.Combine(Server, "Upload")
        If (System.IO.Directory.exists(path)) Then
            path = system.io.path.combine(path, "FCSRun")
            If (System.IO.Directory.exists(path)) Then
                path = system.io.path.combine(path, mUsername)
                If (System.IO.Directory.exists(path)) Then Server_Data = (System.IO.Directory.GetDirectories(path).length > 0)
            End If
        End If

    End Function

    Private Shared mCurRunValid As Boolean = False
    Public Shared ReadOnly Property CurRunValid() As Boolean
        Get
            If (mXML.Exists("CurRun") And mXML.Exists("DataRoot")) Then
                Return (System.IO.Directory.exists(system.io.path.combine(mXML.Exists("DataRoot"), mXML.Exists("CurRun"))))
            Else
                Return False
            End If
        End Get
    End Property

    ' dgp rev 6/19/09 Default Run
    Private Shared Function GetAnyData() As String

        ' read previous setup and validate
        If (Data_Root IsNot Nothing) Then
            If (System.IO.Directory.Exists(mXML.GetSetting("DataRoot"))) Then
                If (System.IO.Directory.GetDirectories(mXML.GetSetting("DataRoot")).Length > 0) Then
                    ' dgp rev 7/16/08 dist root may be emtpy or populated at this point
                    mNoData = False
                    Return System.IO.Directory.GetDirectories(mXML.GetSetting("DataRoot"))(0)
                End If
            Else
                ' dgp rev 7/16/08 no distribution at this point
                Log_Info("DataRoot no longer valid")
            End If
        End If
        ' dgp rev 6/22/09 FlowRoot doesn't even exist, so exit
        If FlowRoot Is Nothing Then Return Nothing

        Dim item
        Dim tmp
        ' dgp rev 6/22/09 find a populated data
        For Each item In RootList
            tmp = System.IO.Path.Combine(item, "Data")
            If (Check_Current(tmp) = mState.Populated) Then
                mXML.PutSetting("DataRoot", tmp)
                Return System.IO.Directory.GetDirectories(tmp)(System.IO.Directory.GetDirectories(tmp).Length - 1)
            End If
        Next

        ' dgp rev 6/22/09 No data locally, so download primer
        If (Check_Data_Server()) Then
            ' dgp rev 7/16/08 if server has a dist, then create folder
            tmp = System.IO.Path.Combine(FlowRoot, "Data")
            If (Utility.Create_Tree(tmp)) Then
                ' dgp rev 7/16/08 download dist, if successful, set current
                If (Download_Data(mServerData, tmp)) Then
                    ' dgp rev 7/16/08 download dist, if successful, set current
                    mXML.PutSetting("DataRoot", tmp)
                    Return System.IO.Directory.GetDirectories(tmp)(0)
                End If
            End If
        End If

        GetAnyData = False
        Log_Info("Failed to establish any data")
        Return Nothing

    End Function

    ' dgp rev 12/5/08 Browse Run instance
    Private Shared mDefaultRun As FCS_Classes.FCSRun
    Public Shared Property DefaultRun() As FCS_Classes.FCSRun
        Get
            If mDefaultRun Is Nothing Then mDefaultRun = New FCSRun(LatestData)
            Return mDefaultRun
        End Get
        Set(ByVal value As FCS_Classes.FCSRun)
            mDefaultRun = value
        End Set
    End Property

    ' dgp rev 12/5/08 Browse Run instance
    Private Shared mBrowseRun As FCS_Classes.FCSRun
    Public Shared Property BrowseRun() As FCS_Classes.FCSRun
        Get
            If mBrowseRun Is Nothing Then mBrowseRun = DefaultRun
            Return mBrowseRun
        End Get
        Set(ByVal value As FCS_Classes.FCSRun)
            mBrowseRun = value
        End Set
    End Property

    ' dgp rev 12/5/08 Retrieve Newly Browse Run
    Public Shared ReadOnly Property RunChange() As Boolean
        Get
            If (mBrowseRun Is Nothing) Then Return False
            Return (Not System.IO.Path.GetDirectoryName(mBrowseRun.Data_Path) = CurRun)
        End Get
    End Property

    ' dgp rev 7/15/08 Retrieve Current Run
    Public Shared Property CurRun() As String
        Get
            If (mCurRun IsNot Nothing) Then Return mCurRun
            ' read previous setup and validate
            If (mXML.Exists("CurRun")) Then
                If (System.IO.Directory.Exists(mXML.GetSetting("CurRun"))) Then
                    Dim test = mXML.GetSetting("CurRun")
                    If (System.IO.Directory.Exists(System.IO.Path.Combine(Data_Root, test))) Then mCurRun = test
                End If
            End If
            If mCurRun Is Nothing Then
                If (RunArray.Count > 0) Then
                    mCurRun = RunArray.Item(0)
                End If
            End If
            Return mCurRun
        End Get
        Set(ByVal value As String)
            mCurRun = value
            mXML.PutSetting("CurRun", value)
        End Set
    End Property

    ' dgp rev 12/5/08 New Data has been added to runarray
    Private Shared mNewDataFlag As Boolean = False
    Public Shared Property NewDataFlag() As Boolean
        Get
            Return mNewDataFlag
        End Get
        Set(ByVal value As Boolean)
            mNewDataFlag = value
        End Set
    End Property

    ' dgp rev 6/4/08 List of local runs in current data area
    Public Shared ReadOnly Property RunArray() As ArrayList
        Get
            If (mRunArray Is Nothing) Then Scan_Runs()
            Return mRunArray
        End Get
    End Property

    ' dgp rev 6/4/08 do runs exists
    Public Shared ReadOnly Property RunsExist() As Boolean
        Get
            Return (RunArray.Count > 0)
        End Get
    End Property

    ' dgp rev 6/4/08 do runs exists
    Public Shared ReadOnly Property RunCount() As Integer
        Get
            Return RunArray.Count
        End Get
    End Property

    Private Shared mScanFlag As Boolean = True

    ' dgp rev 5/22/08 Resort the runs
    Public Shared Sub Sort_Runs()

        '        Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        ' Sorts the values of the ArrayList using the reverse case-insensitive comparer.
        Dim myComparer = New RunCompare
        RunCompare.Reps = 0

        Try
            If (RunDisplay.Count > 1) Then RunDisplay.Sort(myComparer)
        Catch ex As Exception

        End Try

    End Sub

    ' dgp rev 6/4/08 Memebers concerned with Depot Area
    ' dgp rev 8/6/08 the idea here is to include the depot only when needed
    ' and not upon every instance, as with the data, work and distribution areas.
    Private Shared mDepotRoot As String            ' a single Depot Root
    Private Shared mDepotRoots As ArrayList        ' list of multiple Depot Roots
    Private Shared mDepotScan As Boolean = False   ' Depot Roots have been scanned
    Private Shared mDepotState As mState


    ' dgp rev 6/4/08 Memebers concerned with Data Area
    Private Shared mDataRoot As String            ' a single Data Root
    Private Shared mTestDataRoot As String            ' a single Data Root
    Private Shared mDataRoots As ArrayList        ' list of multiple Data Roots
    Private Shared mDataRoot_OTE As Boolean = False   ' Data Roots have been scanned
    Private Shared mOneTimeUserRootCheck As Boolean = False   ' User Roots have been scanned
    Private Shared mOneTimeWorkRootCheck As Boolean = False   ' Work Roots have been scanned
    Private Shared mDataState As mState
    Private Shared mUserState As mState
    Private Shared mCurData As String
    Private Shared mDataChange As String = False
    Private Shared mCurDataValid As Boolean
    Private Shared mLatestData As String

    ' dgp rev 5/29/08 Return the latest run 
    Private Shared mLatestRun As String

    ' dgp rev one-time scan for data
    ' dgp rev 5/28/08 Scan for data in flowroot data area
    ' dgp rev 6/4/08 Assign Data Scan Flag, Data Root Exists Flag,
    ' Data Root List, Data Root and Latest Run
    Public Shared Sub Reset_Data_Scan()

        mDataRoot_OTE = False

    End Sub


    ' dgp rev 11/24/09 Scan roots for local data
    Private Shared Function RemoteDataScan() As Boolean

        ' dgp rev 6/22/09 No data locally, so download primer
        If (Check_Data_Server()) Then
            ' dgp rev 7/16/08 if server has a dist, then create folder
            mTestDataRoot = System.IO.Path.Combine(FlowRoot, "Data")
            If (Utility.Create_Tree(mTestDataRoot)) Then
                ' dgp rev 7/16/08 download dist, if successful, set current
                If (Download_Data(mServerData, mTestDataRoot)) Then
                    ' dgp rev 7/16/08 download dist, if successful, set current
                    mXML.PutSetting("DataRoot", mTestDataRoot)
                    Return True
                End If
            End If
        End If

        Log_Info("Failed to establish any data")
        Return Nothing


    End Function

    ' dgp rev 11/24/09
    Private Shared mSourceMoveData = Nothing

    ' dgp rev 11/24/09 Scan roots for local data
    Private Shared Function LocalDataScan() As Boolean

        Dim path As String
        mSourceMoveData = Nothing

        ' dgp rev 11/24/09 Scan locally for data
        Dim item
        For Each item In RootList
            path = System.IO.Path.Combine(item, "Data")
            If (System.IO.Directory.Exists(path)) Then
                ' dgp rev 5/28/08 make sure data root isn't empty
                If (System.IO.Directory.GetDirectories(path).Length > 0) Then mDataRoots.Add(path)
            End If
        Next

        If ((mDataRoots.Count > 0)) Then
            ' dgp rev 6/4/08 use property so data is saved to XML
            mSourceMoveData = mDataRoots.Item(mDataRoots.Count - 1)
        End If
        Return (mSourceMoveData IsNot Nothing)

    End Function

    ' dgp rev 11/24/09 Move target data into Flow Root
    Private Shared Function MoveData() As Boolean

        Return True

    End Function

    Private Shared mDataRootExists As Boolean = True
    Private Shared ReadOnly Property DataRootExists As Boolean
        Get
            If mDataRoot_OTE Then EstablishDataRoots()
            Return mDataRootExists
        End Get
    End Property

    ' dgp rev 7/10/08 Scan roots for a data area with containing data
    ' dgp rev 11/24/09 Scan roots for local data
    Private Shared Sub EstablishDataRoots()

        mDataRoot_OTE = True
        mDataRoots = New ArrayList

        Dim FoundData = False

        mDataRoot = Nothing
        ' dgp rev 11/24/09 look for persistant data settings
        If (PersistantDataRoot()) Then
            mDataRoot = mTestDataRoot
            mDataRootExists = True
            Return
        Else
            ' dgp rev 11/24/09 look for data in FlowRoot
            If FlowRoot Is Nothing Then Return
            mTestDataRoot = System.IO.Path.Combine(FlowRoot, "Data")
            If (System.IO.Directory.Exists(mTestDataRoot)) Then
                If System.IO.Directory.GetDirectories(mTestDataRoot).Length > 0 Then
                    mDataRoot = mTestDataRoot
                    mDataRootExists = True
                    Return
                End If
            End If
        End If

        ' dgp rev 11/24/09 No Root data, so scan for any other data to access
        If Not LocalDataScan() Then
            If Not RemoteDataScan() Then
                mDataRootExists = Utility.Create_Tree(mTestDataRoot)
                Return
            End If
        End If

        ' dgp rev 11/24/09 Move the non-root data
        MoveData()

    End Sub

    ' dgp rev 7/10/08 Scan roots for a data area with containing data
    ' dgp rev 11/24/09 Scan roots for local data
    Private Shared Sub EstablishDataRoot()

        mDataRoot_OTE = True
        mDataRoots = New ArrayList

        Dim FoundData = False

        mDataRoot = Nothing
        ' dgp rev 11/24/09 look for persistant data settings
        If (PersistantDataRoot()) Then
            mDataRoot = mTestDataRoot
            Return
        Else
            ' dgp rev 11/24/09 look for data in FlowRoot
            If FlowRoot Is Nothing Then Return
            mTestDataRoot = System.IO.Path.Combine(FlowRoot, "Data")
            If (System.IO.Directory.Exists(mTestDataRoot)) Then
                If System.IO.Directory.GetDirectories(mTestDataRoot).Length > 0 Then
                    mDataRoot = mTestDataRoot
                    Return
                End If
            End If
        End If

        ' dgp rev 11/24/09 No Root data, so scan for any other data to access
        If Not LocalDataScan() Then
            If Not RemoteDataScan() Then
                Utility.Create_Tree(mTestDataRoot)
                Return
            End If
        End If

        ' dgp rev 11/24/09 Move the non-root data
        MoveData()

    End Sub

    ' dgp rev 12/2/09 
    Private Shared Function GlobalUserTransfer() As Boolean


        Return True

    End Function

    Private Shared mUserRoots
    Private Shared mTestUserRoot

    ' dgp rev 7/10/08 Scan roots for a data area with containing data
    ' dgp rev 11/24/09 Scan roots for local data
    ' dgp rev 12/2/09 
    Private Shared Sub EstablishUserRoot()

        mOneTimeUserRootCheck = True
        mUserRoots = New ArrayList
        mNoUser = True

        Dim FoundUser = False

        mUserRoot = Nothing
        ' dgp rev 11/24/09 look for persistant User settings
        If (PersistantUserRoot()) Then
            mUserRoot = mTestUserRoot
        Else
            ' dgp rev 11/24/09 look for User in FlowRoot
            If FlowRoot Is Nothing Then Return
            mTestUserRoot = System.IO.Path.Combine(FlowRoot, "Users")
            mTestUserRoot = System.IO.Path.Combine(mTestUserRoot, mUsername)
            If (Utility.Create_Tree(mTestUserRoot)) Then
                mUserRoot = mTestUserRoot
            End If
        End If

        mNoUser = mUserRoot Is Nothing
        Return

    End Sub

    ' dgp rev 7/10/08 Scan roots for a data area with containing data
    Private Shared Sub ScanDepot()

        mDepotScan = True
        mDepotRoots = New ArrayList

        Dim path As String

        Dim item
        For Each item In RootList
            path = System.IO.Path.Combine(item, "Depot")
            If (System.IO.Directory.Exists(path)) Then
                ' dgp rev 5/28/08 make sure Depot root isn't empty
                If (System.IO.Directory.GetDirectories(path).Length > 0) Then mDepotRoots.Add(path)
            End If
        Next

        If ((mDepotRoots.Count > 0)) Then
            ' dgp rev 6/4/08 use property so Depot is saved to XML
            Depot_Root = mDepotRoots.Item(mDepotRoots.Count - 1)
        End If

    End Sub

    ' dgp rev 6/4/08
    Public Shared ReadOnly Property DepotRoots() As ArrayList
        Get
            If (Not mDepotScan) Then ScanDepot()
            Return mDepotRoots
        End Get
    End Property

    ' dgp rev 6/4/08
    Public Shared ReadOnly Property DataRoots() As ArrayList
        Get
            If (Not mDataRoot_OTE) Then EstablishDataRoot()
            Return mDataRoots
        End Get
    End Property

    ' dgp rev 7/9/08 Does work exist, 
    Public Shared ReadOnly Property WorkExists() As Boolean
        Get
            If (mWorkState = mState.None) Then Return False
            Return (System.IO.Directory.GetDirectories(mWorkRoot).Length > 0)
        End Get
    End Property

    Private Shared mServerChksum_Root
    Private Shared mServerChksumState
    Private Shared mNoServerChksum
    Private Shared mServerChksumScan
    Private Shared mServerChksumRoots

    ' dgp rev 7/10/08 Scan roots for a data area with containing data
    Private Shared Sub ScanServerChksum()

        mServerChksumScan = True
        mServerChksumRoots = New ArrayList

        Dim path As String

        Dim item
        For Each item In RootList
            path = System.IO.Path.Combine(item, "Server")
            If (System.IO.Directory.Exists(path)) Then
                ' dgp rev 5/28/08 make sure Depot root isn't empty
                If (System.IO.Directory.GetDirectories(path).Length > 0) Then mServerChksumRoots.Add(path)
            End If
        Next

        If ((mServerChksumRoots.Count > 0)) Then
            ' dgp rev 6/4/08 use property so Depot is saved to XML
            ServerChksum_Root = mServerChksumRoots.Item(mServerChksumRoots.Count - 1)
        Else
            path = System.IO.Path.Combine(FlowRoot, "Server")
            ' dgp rev 7/16/08 if server has a dist, then create folder
            If (Utility.Create_Tree(path)) Then
                ServerChksum_Root = path
            End If
        End If

    End Sub

    ' dgp rev 6/4/08 return the data root
    Public Shared Property ServerChksum_Root() As String
        Get
            If (Not mServerChksumScan) Then ScanServerChksum()
            Return mServerChksum_Root
        End Get
        ' dgp rev 6/4/08 change Depot root in XML also
        Set(ByVal value As String)
            mServerChksum_Root = value
            ' dgp rev 7/16/08 determine state none, empty or populated
            mServerChksumState = Check_Current(value)
            mNoServerChksum = (Not mServerChksumState = mState.Populated)
            ' dgp rev 7/16/08 persistence
            mXML.PutSetting("ServerChksumRoot", value)
        End Set

    End Property

    ' dgp rev 6/4/08 return the data root
    Public Shared Property Depot_Root() As String
        Get
            If (Not mDepotScan) Then ScanDepot()
            Return mDepotRoot
        End Get
        ' dgp rev 6/4/08 change Depot root in XML also
        Set(ByVal value As String)
            mDepotRoot = value
            ' dgp rev 7/16/08 determine state none, empty or populated
            mDepotState = Check_Current(value)
            mNoDepot = (Not mDepotState = mState.Populated)
            ' dgp rev 7/16/08 persistence
            mXML.PutSetting("DepotRoot", value)
        End Set

    End Property

    ' dgp rev 6/4/08 return the data root
    Public Shared Property Data_Root() As String
        Get
            If (Not mDataRoot_OTE) Then EstablishDataRoots()
            Return mDataRoot
        End Get
        ' dgp rev 6/4/08 change data root in XML also
        Set(ByVal value As String)
            mDataRoot = value
            ' dgp rev 7/16/08 determine state none, empty or populated
            mDataState = Check_Current(value)
            mNoData = (Not mDataState = mState.Populated)
            ' dgp rev 7/16/08 persistence
            mXML.PutSetting("DataRoot", value)
        End Set

    End Property

    ' dgp rev 6/4/08 return the data root
    ' dgp rev 12/2/09 
    Public Shared Property User_Root() As String
        Get
            If (Not mOneTimeUserRootCheck) Then EstablishUserRoot()
            Return mUserRoot
        End Get
        ' dgp rev 6/4/08 change User root in XML also
        Set(ByVal value As String)
            mUserRoot = value
            ' dgp rev 7/16/08 determine state none, empty or populated
            mUserState = Check_Current(value)
            mNoUser = (Not mUserState = mState.Populated)
            ' dgp rev 7/16/08 persistence
            mXML.PutSetting("UserRoot", value)
        End Set

    End Property

    ' a single Data Root
    ' dgp rev 6/4/08 Retrieve the lastest local run
    Public Shared ReadOnly Property LatestData() As String
        Get
            If (Not mDataRoot_OTE) Then EstablishDataRoot()
            mLatestData = GDR_Latest()
            Return mLatestData
        End Get
    End Property

    ' dgp rev 6/4/08 Reset runs 
    Public Shared Sub Reset_Runs()

        mRunArray = Nothing

    End Sub

    Private mRemoveEmpty As Boolean = True
    ' dgp rev 12/5/07 Rescan local runs and sort the list, try to do this only once
    ' unless new data is added or removed
    Private Shared Sub Scan_Runs()

        mRunArray = New ArrayList
        If (System.IO.Directory.Exists(Data_Root)) Then
            Dim item
            For Each item In System.IO.Directory.GetDirectories(mDataRoot)
                RaiseEvent FSRemoteEvent(System.IO.Path.GetFileName(item) + "...")
                mRunArray.Add(System.IO.Path.GetFileName(item))
            Next
            ' dgp rev 6/5/08 always sort the newly scanned runs
            If (mRunArray.Count > 1) Then Sort_Runs()
        End If

    End Sub

    ' dgp rev 2/22/08 Concatenate the Root on Remote Server
    Public Shared ReadOnly Property RemoteRoot() As String
        Get
            If (FlowServer.Server_Up) Then
                Return String.Format("\\{0}\{1}\Users\{2}", FlowServer.FlowServer, mShare_Flow, mUsername)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Private Shared mSettings = Nothing
    ' dgp rev 2/22/08 Concatenate the Root on Remote Server
    Public Shared Property Settings() As String
        Get
            If mSettings Is Nothing Then mSettings = Establish_Settings()
            Return mSettings
        End Get
        Set(ByVal value As String)
            mSettings = value
        End Set
    End Property

    ' dgp rev 7/9/08 Scan for particular user area
    Public Shared Sub ScanUser()

        Dim Part As String
        Dim item
        ' dgp rev 2/22/08 personal work
        mUserFlag = False
        For Each item In RootList
            Part = System.IO.Path.Combine(System.IO.Path.Combine(item, "Users"), mUsername)
            ' dgp rev 5/23/08 does not guarantee user area
            If (System.IO.Directory.Exists(Part)) Then
                mUserRoot = Part
                mUserFlag = True
            End If
        Next

    End Sub

    ' dgp rev one-time scan for work
    Public Shared Sub ScanWork()

        mWorkScan = True
        mWorkRoots = New ArrayList
        Dim Part As String

        Dim Work_List As New ArrayList

        ' dgp rev 2/22/08 generic work
        Dim item

        ' dgp rev 6/18/09 problem occurs with empty Work folder
        For Each item In RootList
            Part = System.IO.Path.Combine(item, "Work")
            If (System.IO.Directory.Exists(Part)) Then
                If (System.IO.Directory.GetDirectories(Part).Length > 0) Then Work_List.Add(System.IO.Directory.GetDirectories(Part))
            End If
        Next

        ' dgp rev 2/22/08 personal work
        For Each item In RootList
            Part = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.Combine(item, "Users"), mUsername), "Work")
            ' dgp rev 5/23/08 does not guarantee projects or sessions
            If (System.IO.Directory.Exists(Part)) Then
                If (System.IO.Directory.GetDirectories(Part).Length > 0) Then Work_List.Add(System.IO.Directory.GetDirectories(Part))
            End If
        Next

        Dim base, proj As Object

        For Each base In Work_List
            If (System.IO.Directory.GetDirectories(base).Length > 0) Then
                'If (System.IO.Directory.GetDirectories(path).length > 0) Then
                For Each proj In System.IO.Directory.GetDirectories(base)
                    '                        For Each proj In base.SubFolders
                    If (proj.length > 0) Then
                        mWorkRoots.Add(base.Path.ToString)
                        Exit For
                    End If
                Next
            End If
        Next
'        If (mWorkRoots.Count > 0) Then Work_Root = mWorkRoots.Item(mWorkRoots.Count - 1)

    End Sub


    ' dgp rev one-time scan for a distribution
    Private Shared Sub ScanDist()

        Dim flowitem

        Dim tst = System.IO.Path.Combine(FlowRoot, "Distribution")
        If (System.IO.Directory.Exists(tst)) Then
            mDistRoot = tst
        Else
            For Each flowitem In RootList
                tst = System.IO.Path.Combine(flowitem, "Distribution")
                If (System.IO.Directory.Exists(tst)) Then mDistRoot = tst
            Next
        End If
        If (mDistRoot Is Nothing) Then
            mDistRoot = System.IO.Path.Combine(FlowRoot, "Distribution")
            Utility.Create_Tree(mDistRoot)
        End If

    End Sub

    ' dgp rev 3/29/2011 does a work root need to be created
    Public Shared ReadOnly Property WorkRootExists As Boolean
        Get
            If Work_Root Is Nothing Then Return False
            Return Directory.Exists(Work_Root)
        End Get
    End Property

    ' dgp rev 3/29/2011 is a work merge reuired
    Public Shared ReadOnly Property WorkIsPersonal As Boolean
        Get
            If Work_Root Is Nothing Then Return False
            Return Work_Root.ToLower.Contains("users") And Work_Root.ToLower.Contains(mUsername)
        End Get
    End Property

    Private Shared mAllFiles() As FileInfo

    ' dgp rev 5/12/2011 Scan files for latest work
    Private Shared Function ScanWorkRoot() As FileInfo()

        Dim dirinfo As DirectoryInfo = New DirectoryInfo(PersonalWorkRoot)
        '        mAllFiles = dirinfo.GetFiles("*.xml", SearchOption.TopDirectoryOnly)
        mAllFiles = dirinfo.GetFiles("fcs_files.lis", SearchOption.AllDirectories)
        Array.Sort(mAllFiles, New ModDateCompare)
        Return mAllFiles

    End Function


    ' dgp rev 5/11/2011 XML Document for PV-Wave
    Private Shared mPVW_XML_Doc As XmlDocument

    ' dgp rev 5/11/2011 get the XML config file 
    Public Shared Function PVW_XML_Hash() As Hashtable

        PVW_XML_Hash = New Hashtable

        ' dgp rev 5/1/07 Dynamic settings
        '        dyn = New Dynamic(system.io.path.combine(user_path, "Dynamic.XML"))
        If (Not System.IO.File.Exists(FlowStructure.PVW_XML_File)) Then Exit Function
        ' dgp rev 3/31/2011 add filename key to attribute list
        '            If Not PVW_from_XML.ContainsKey("Filename") Then PVW_from_XML.Add("Filename", FlowStructure.PVW_XML_File)
        '            Log_Info("Reading XML " + FlowStructure.PVW_XML_File)
        mPVW_XML_Doc = New XmlDocument
        Dim XMLNode As XmlNode
        mPVW_XML_Doc.Load(FlowStructure.PVW_XML_File)
        XMLNode = mPVW_XML_Doc.SelectSingleNode("Environment/FlowControl/Settings")
        If (XMLNode Is Nothing) Then
            '                PVW_from_XML.Add("Error", True)
        Else
            Dim idx
            For idx = 0 To XMLNode.Attributes.Count - 1
                PVW_XML_Hash.Add(XMLNode.Attributes.Item(idx).Name, XMLNode.Attributes.Item(idx).Value)
                '                    Log_Info(XMLNode.Attributes.Item(idx).Name + " => " + XMLNode.Attributes.Item(idx).Value)
            Next
        End If

    End Function

    ' dgp rev 5/12/2011 Establish last work 
    Private Shared Sub Establish_LastWork()

        mLastWork_OTE = True

        ' dgp rev 5/12/2011 use persistent data for efficiency
        If (PVW_XML_exists) Then
            ' dgp rev 5/12/2011 check that it contains a work path
            If PVW_XML_Hash.Contains("work_path") Then
                mLastWork = PVW_XML_Hash("work_path")
                mLastWorkValid = ValidWork(mLastWork)
            End If
        End If

        ' dgp rev 5/12/2011 if persistent last work is not available then scan
        If Not mLastWorkValid Then
            If PersonalWorkRootExists Then
                Dim path As FileInfo
                If System.IO.Directory.Exists(Data_Root) Then
                    For Each path In ScanWorkRoot()
                        If ValidWork(path.Directory.ToString) Then
                            mLastWork = path.Directory.ToString
                            mLastWorkValid = True
                            Exit For
                        End If
                    Next
                End If
            End If
        End If

    End Sub

    Private Shared mLastWork_OTE As Boolean = False
    Private Shared ReadOnly Property LastWork As String
        Get
            If Not mLastWork_OTE Then Establish_LastWork()
            Return mLastWork
        End Get
    End Property

    ' dgp rev 5/12/2011 
    Private Shared Function LastWorkExists() As Boolean

        ' dgp rev 5/12/2011 check for the XML file
        If Not mLastWork_OTE Then Establish_LastWork()
        Return mLastWorkValid

    End Function

    ' dgp rev 5/11/2011 Well Structured Work, no generic, only personal
    Public Shared ReadOnly Property WorkWellStructured As Boolean
        Get
            Return LastWorkExists()
        End Get
    End Property

    ' dgp rev 6/23/09 Get the Work Root if it exists
    Private Shared Function FindWorkRoot() As String

        If FlowRoot Is Nothing Then Return Nothing

        If (System.IO.Directory.Exists(System.IO.Path.Combine(System.IO.Path.Combine(UsersRoot, mUsername), "Work"))) Then
            Return System.IO.Path.Combine(System.IO.Path.Combine(UsersRoot, mUsername), "Work")
        End If
        If (System.IO.Directory.Exists(System.IO.Path.Combine(FlowRoot, "Work"))) Then
            Return System.IO.Path.Combine(FlowRoot, "Work")
        End If
        Return System.IO.Path.Combine(System.IO.Path.Combine(UsersRoot, mUsername), "Work")

    End Function

    Private Shared mWorkRoot_OTE As Boolean = False
    Private Shared mWorkRootExists As Boolean = False
    Private Shared mWorkRoot As String

    ' dgp rev 2/22/08 The Work root
    Public Shared ReadOnly Property Work_Root()
        Get
            ' dgp rev 5/12/2011 if valid current work, then valid work root
            If CurWorkValid Then
                If mCurWork.ToLower.Contains("users") And mCurWork.ToLower.Contains(mUsername) Then
                    Return Directory.GetParent(Directory.GetParent(mCurWork).ToString).ToString
                End If
            End If

            ' dgp rev 5/12/2011 if current is not valid, may return an empty work root
            If Not mWorkRoot_OTE Then mWorkRoot = Establish_WorkRoot()
            Return mWorkRoot

        End Get
    End Property

    ' dgp rev 5/23/08 Work roots available
    Public Shared Property SettingsRoots() As ArrayList
        Get
            Return mSettingsRoots
        End Get
        Set(ByVal value As ArrayList)
            mSettingsRoots = value
        End Set
    End Property

    ' dgp rev 5/23/08 Work roots available
    Public Shared Property WorkRoots() As ArrayList
        Get
            Return mWorkRoots
        End Get
        Set(ByVal value As ArrayList)
            mWorkRoots = value
        End Set
    End Property

    ' dgp rev 7/16/08 Distribution root for local revisions
    Public Shared ReadOnly Property Dist_Root() As String
        Get
            If (mDistRoot Is Nothing) Then ScanDist()
            Return mDistRoot
        End Get
    End Property

    ' dgp rev 7/20/07 
    Public Shared ReadOnly Property RootList() As ArrayList
        Get
            If (mRoots Is Nothing) Then ScanRoot()
            Return mRoots
        End Get
    End Property

    ' dgp rev 6/22/09 Get the current FlowRoot
    Public Shared ReadOnly Property FlowRoot() As String
        Get
            If mFlowRoot Is Nothing Then Establish_Root()
            Return mFlowRoot
        End Get
    End Property

    ' dgp rev 2/22/08 hold persistant data in XML file
    Private Shared mXML As New Dynamic("FlowStructure")

    ' dgp rev 2/22/08 error messages 
    Private Shared mMessage As String

    ' dgp rev 2/22/08 error messages 
    Private Shared mSep As Char = System.IO.Path.DirectorySeparatorChar

    ' dgp rev 2/22/08 One-time scan to fill in values
    Private Shared mDrives As ArrayList
    Private Shared mRoots As ArrayList

    ' dgp rev 7/20/07 
    Private Shared Sub ScanRoot()

        Dim tstDrv As String
        mRoots = New ArrayList
        mDrives = HelperClasses.Utility.LocalDrives
        ' dgp rev 2/19/08 scan drives for valid locations
        Dim drv

        For Each drv In HelperClasses.Utility.LocalDrives
            tstDrv = drv + "FlowRoot"
            If System.IO.Directory.Exists(tstDrv) Then mRoots.Add(tstDrv)
        Next

    End Sub

    ' dgp rev 7/14/08 Create the initial root
    Private Shared Function Create_Root() As Boolean



    End Function

    ' dgp rev 7/10/08 Check path characteristics
    Private Shared Function Check_Current(ByVal path As String) As mState

        RaiseEvent FSRemoteEvent("Check " + path + "...")
        Check_Current = mState.None
        If (System.IO.Directory.Exists(path)) Then
            Check_Current = mState.Empty
            If (System.IO.Directory.GetDirectories(path).Length > 0) _
              Or (System.IO.Directory.GetFiles(path).Length > 0) Then Check_Current = mState.Populated
        End If

    End Function

    ' dgp rev 7/10/08 Check path characteristics
    Private Shared Function Check_Path(ByVal path As String) As mFolderFlag

        Check_Path = mFolderFlag.None
        If (System.IO.Directory.Exists(path)) Then
            If (System.IO.Directory.GetDirectories(path).Length > 0) Then Check_Path = mFolderFlag.Folders
            If (System.IO.Directory.GetFiles(path).Length > 0) Then Check_Path = Check_Path Or mFolderFlag.Files
            If (Check_Path = mFolderFlag.None) Then Check_Path = mFolderFlag.Empty
        End If

    End Function

    ' dgp rev 2/20/08 Does Root Exist
    Private Shared Function Check_Root() As Boolean

        Check_Root = False
        If (mRoots.Count > 0) Then
            Check_Root = True
            mFlowRoot = mRoots.Item(mRoots.Count - 1)
        End If

    End Function

    ' dgp rev 2/22/08 Set the FlowRoot
    Public Shared Function Set_Root(ByVal path As String) As Boolean

        Set_Root = False
        If (mRoots.ToString.ToLower.Contains(path.ToLower)) Then
            mFlowRoot = path
            Set_Root = True
        End If

    End Function

    ' dgp rev 2/20/08 Create the FlowRoot
    Private Shared Function Make_Root() As Boolean

        Make_Root = False
        mRoots = New ArrayList
        ' dgp rev 2/19/08 review valid locations and select or create FlowRoot
        If (mDrives.Count > 0) Then
            Dim tst As String
            ' dgp rev 6/15/09 
            tst = System.IO.Path.Combine(Utility.LargestDrive, "FlowRoot")
            If (Utility.Create_Tree(tst)) Then
                mFlowRoot = tst
                mRoots.Add(tst)
                Make_Root = True
            Else
                mMessage = "Root Creation Failed - " + tst
            End If
        Else
            mMessage = "No valid local drives"
        End If

    End Function


    ' dgp rev 2/19/08 First time running FlowRoot, must initialize
    Public Shared Function Init_Root() As Boolean

        Init_Root = False

        ' dgp rev 7/20/07 
        ScanRoot()
        If (Check_Root()) Then
            Init_Root = True
            mNoRoot = False
        Else
            Init_Mask()
            If (Make_Root()) Then
                mNoRoot = False
                Init_Root = True
            End If
        End If

    End Function

    Private mForceWork = Nothing

    ' dgp rev 8/6/08 First check to see if depot was already established 
    ' and save to XML
    Private Shared Function GCR_XML() As Boolean

        mDepotState = mState.None
        ' dgp rev 2/20/08 keep this simple, if XML settings exists, then read it.
        If (mXML.NewFlag) Then Return False

        ' read previous setup and validate
        If (mXML.Exists("DepotRoot")) Then
            mDepotState = Check_Current(mXML.GetSetting("DepotRoot"))
            If (System.IO.Directory.Exists(mXML.GetSetting("DepotRoot"))) Then
                Depot_Root = mXML.GetSetting("DepotRoot")
                ' dgp rev 7/16/08 Depot root may be emtpy or populated at this point
                mNoDepot = (Not mDepotState = mState.Populated)
                Return True
            Else
                ' dgp rev 7/16/08 no Depot at this point
                mNoDepot = True
                Log_Info("DepotRoot no longer valid")
                Return False
            End If
        End If

    End Function

    ' dgp rev 2/22/08 Get XML Data Path
    Private Shared Function GDR_XML() As Boolean

        mDataState = mState.None
        ' dgp rev 2/20/08 keep this simple, if XML settings exists, then read it.
        If (mXML.NewFlag) Then Return False

        ' read previous setup and validate
        If (mXML.Exists("DataRoot")) Then
            mDataState = Check_Current(mXML.GetSetting("DataRoot"))
            If (System.IO.Directory.Exists(mXML.GetSetting("DataRoot"))) Then
                Data_Root = mXML.GetSetting("DataRoot")
                ' dgp rev 7/16/08 Data root may be emtpy or populated at this point
                mNoData = (Not mDataState = mState.Populated)
                Return Not mNoData
            Else
                ' dgp rev 7/16/08 no Data at this point
                mNoData = True
                Log_Info("DataRoot no longer valid")
                Return False
            End If
        End If

    End Function

    ' dgp rev 2/22/08 Get XML Data Path
    Private Shared Function GWR_FlowRoot() As Boolean

        If FlowRoot Is Nothing Then Return False
        If (Not System.IO.Directory.Exists(System.IO.Path.Combine(System.IO.Path.Combine(UsersRoot, mUsername), "Work"))) Then Return False
        '        Work_Root = System.IO.Path.Combine(System.IO.Path.Combine(UsersRoot, mUsername), "Work")
        Return True

    End Function

    ' dgp rev 2/22/08 Check the current FlowRoot
    Private Shared Function GPR_FlowRoot() As Boolean

        If (Not System.IO.Directory.Exists(System.IO.Path.Combine(mFlowRoot, "Distribution"))) Then Return False
        Return True

    End Function

    ' dgp rev 8/6/08 Check the current root
    Private Shared Function GCR_FlowRoot() As Boolean

        ' dgp rev 6/4/08 only set the Depot root from possibilities in Depot root list
        If (Not System.IO.Directory.Exists(System.IO.Path.Combine(mFlowRoot, "Depot"))) Then Return False
        Depot_Root = System.IO.Path.Combine(mFlowRoot, "Depot")
        Return True

    End Function

    ' dgp rev 2/22/08 Check Flow Root for Data Area
    Private Shared Function GDR_FlowRoot() As Boolean

        ' dgp rev 6/4/08 only set the data root from possibilities in data root list
        If (Not System.IO.Directory.Exists(System.IO.Path.Combine(mFlowRoot, "Data"))) Then Return False
        Data_Root = System.IO.Path.Combine(mFlowRoot, "Data")
        ' dgp rev 6/18/09 Data Root may exists, but does it hold anydata
        Return System.IO.Directory.GetDirectories(Data_Root).Length > 0

    End Function

    ' dgp rev 2/22/08 Get XML Work Path
    Private Shared Function GWR_Scan() As Boolean

        ScanWork()
        If (mWorkRoots.Count = 0) Then Return False
        '        Work_Root = mWorkRoots.Item(mWorkRoots.Count - 1)

    End Function

    ' dgp rev 8/6/08 Scan a know roots for depot
    Private Shared Function GCR_Scan() As Boolean

        ScanDepot()
        If (mDepotRoots.Count = 0) Then Return False
        ' dgp rev 6/4/08 use property so path is saved to XML
        Depot_Root = mDepotRoots.Item(mDepotRoots.Count - 1)
        Return True

    End Function

    ' dgp rev 2/22/08 Get XML Data Path
    Private Shared Function GDR_Scan() As Boolean

        EstablishDataRoot()
        If (mDataRoots.Count = 0) Then Return False
        ' dgp rev 6/4/08 use property so path is saved to XML
        Data_Root = mDataRoots.Item(mDataRoots.Count - 1)
        Return True

    End Function
    ' dgp rev 5/29/08 Retrieve the most recent data run
    Public Shared Function GDR_Latest() As String

        GDR_Latest = ""
        Dim run
        Dim last_date As DateTime
        Dim last_run As String = ""

        For Each run In System.IO.Directory.GetDirectories(Data_Root)
            If (System.IO.File.GetCreationTime(run).CompareTo(last_date) = 1) Then
                last_date = System.IO.File.GetCreationTime(run)
                last_run = run
            End If
        Next
        GDR_Latest = last_run

    End Function

    ' dgp rev 4/12/07 Create a Unique Name
    Public Shared Function Unique_Name() As String

        Return Format(Now(), "yyyyMMddhhmmss")

    End Function

    ' dgp rev 6/5/07 create of list of the data files
    ' dgp rev 9/28/2010 replace the hard code with a class reference to the fcs file list
    Private Shared Sub Create_Data_List(ByVal session, ByVal data)

        If data Is Nothing Then Return

        Dim objList As FCS_List = New FCS_List(session)
        Dim sw As New StreamWriter(objList.List_Spec)

        Dim item
        For Each item In System.IO.Directory.GetFiles(data)
            sw.WriteLine(item.ToString)
        Next
        sw.Close()

    End Sub

    Private Shared mServerData
    ' dgp rev 7/15/08 Download the latest data from the Server
    Private Shared Function Check_Data_Server() As Boolean

        Try

            Dim Server = String.Format("\\{0}", mServer)

            Check_Data_Server = False
            Dim path = System.IO.Path.Combine(Server, "Upload")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "FCSRun")
                If (System.IO.Directory.Exists(path)) Then
                    path = System.IO.Path.Combine(path, mUsername)
                    If (System.IO.Directory.Exists(path)) Then
                        Dim run
                        If (System.IO.Directory.GetDirectories(path).Length > 0) Then
                            For Each run In System.IO.Directory.GetDirectories(path)
                                If (System.IO.Directory.GetFiles(run).Length > 0) Then
                                    mServerData = run
                                    Check_Data_Server = True
                                    Return True
                                End If
                            Next
                        End If
                    End If

                End If
            End If
            path = System.IO.Path.Combine(Server, "Upload")
            Dim user
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "FCSRun")
                If (System.IO.Directory.Exists(path)) Then
                    For Each user In System.IO.Directory.GetDirectories(path)
                        path = user
                        If (System.IO.Directory.Exists(path)) Then
                            Dim run
                            If (System.IO.Directory.GetDirectories(path).Length > 0) Then
                                For Each run In System.IO.Directory.GetDirectories(path)
                                    If (System.IO.Directory.GetFiles(run).Length > 0) Then
                                        mServerData = run
                                        Check_Data_Server = True
                                        Return True
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If

        Catch ex As Exception
        End Try
        Return False

    End Function


    ' dgp rev 7/15/08 Download the latest distribution from the Server
    Private Shared Function Check_Server() As Boolean

        Dim Server = String.Format("\\{0}", mServer)

        Check_Server = False
        Dim path = System.IO.Path.Combine(Server, "Distribution")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "Versions")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "Current")
                Check_Server = (System.IO.Directory.Exists(path))
            End If
        End If

    End Function

    ' dgp rev 7/15/08 Download the latest distribution
    Private Shared Function Download_Data(ByVal source, ByVal target) As Boolean

        Download_Data = True
        Try
            Utility.DirectoryCopy(source, System.IO.Path.Combine(target, System.IO.Path.GetFileName(source)))
        Catch ex As Exception
            Download_Data = False
        End Try

    End Function

    ' dgp rev 12/2/09 Find Any Valid Work under a given work structure
    Private Shared Function AnyValidWork(ByVal path) As Boolean

        AnyValidWork = False
        If (System.IO.Directory.Exists(path)) Then
            If (System.IO.Directory.GetDirectories(path).Length > 0) Then
                Dim proj
                For Each proj In System.IO.Directory.GetDirectories(path)
                    Dim sess
                    For Each sess In System.IO.Directory.GetDirectories(proj)
                        If (ValidWork(sess)) Then
                            mAnyWorkPath = sess
                            AnyValidWork = True
                            Return True
                        End If
                    Next
                Next
            End If
        End If
        Return False

    End Function

    ' dgp rev 12/2/09 Validate the work path for valid data
    Public Shared Function FCSValidDataList(ByVal path) As ArrayList

        FCSValidDataList = New ArrayList
        ' dgp rev 9/29/2010 Use the FCS List object and not hard code
        Dim objFCSList As FCS_List = New FCS_List(path)
        Dim FileList = objFCSList.List_Spec
        If (System.IO.File.Exists(FileList)) Then

            Dim sr As New StreamReader(FileList.ToString)
            While (Not sr.EndOfStream)
                FCSValidDataList.Add(System.IO.File.Exists(sr.ReadLine).ToString)
            End While
            sr.Close()

        End If

    End Function

    ' dgp rev 12/2/09 Validate the work path for valid data
    Public Shared Function FCSFileList(ByVal path) As ArrayList

        FCSFileList = New ArrayList
        ' dgp rev 9/29/2010 Use the FCS List object and not hard code
        Dim objFCSList As FCS_List = New FCS_List(path)
        Dim FileList = objFCSList.List_Spec
        If (System.IO.File.Exists(FileList)) Then

            Dim sr As New StreamReader(FileList.ToString)
            While (Not sr.EndOfStream)
                FCSFileList.Add(sr.ReadLine)
            End While
            sr.Close()

        End If

    End Function

    ' dgp rev 12/2/09 Validate the work path for valid data
    Public Shared Function FCSListExists(ByVal path) As Boolean

        ' dgp rev 9/29/2010 Use the FCS List object and not hard code
        Dim objFCSList As FCS_List = New FCS_List(path)
        Dim FileList = objFCSList.List_Spec
        Return (System.IO.File.Exists(FileList))

    End Function

    Public Class ModDateCompare

        Implements IComparer

        ' Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare

            Return ModDateCompare(x, y)

        End Function

        Private Function ModDateCompare(ByVal x As Object, ByVal y As Object) As Integer

            Dim File1 As FileInfo
            Dim File2 As FileInfo
            File1 = DirectCast(x, FileInfo)
            File2 = DirectCast(y, FileInfo)
            ModDateCompare = DateTime.Compare(File2.LastWriteTime, File1.LastWriteTime)
        End Function

    End Class


    ' dgp rev 12/2/09 Validate the work path for valid data
    Private Shared Function ValidWork(ByVal path) As Boolean

        ValidWork = False
        ' dgp rev 9/29/2010 Use the FCS List object and not hard code
        Dim objFCSList As FCS_List = New FCS_List(path)
        Dim FileList = objFCSList.List_Spec
        If (System.IO.File.Exists(FileList)) Then

            Dim sr As New StreamReader(FileList.ToString)
            While (Not sr.EndOfStream)
                If (System.IO.File.Exists(sr.ReadLine)) Then
                    ValidWork = True
                    Exit While
                End If
            End While
            sr.Close()

        End If

    End Function

    ' dgp rev 5/29/08 Work Validation
    Private Shared Function Check_Work(ByVal path) As Boolean

        Check_Work = False
        ' dgp rev 5/29/08 is work path valid
        If (System.IO.File.Exists(path)) Then
            Dim sr As New StreamReader(path.ToString)
            While (Not sr.EndOfStream)
                If (System.IO.File.Exists(sr.ReadLine)) Then
                    Check_Work = True
                    Exit While
                End If
            End While
            sr.Close()
        End If

    End Function


    ' dgp rev 6/26/09
    Private Shared mAnyWorkPath As String
    Private Shared mAnyWorkAvailable = Nothing
    Public Shared ReadOnly Property AnyWorkAvailable() As Boolean
        Get
            If (mAnyWorkAvailable Is Nothing) Then mAnyWorkAvailable = Any_Work()
            Return mAnyWorkAvailable
        End Get
    End Property

    ' dgp rev 6/18/09
    Private Shared Function Any_Work() As Boolean

        Dim arr = DistributionServer.ServerDistList()
        ' dgp rev 9/29/2010 Use the FCS List object and not hard code
        Dim objFCSList As FCS_List

        Dim root
        Any_Work = False
        For Each root In FCS_Classes.FlowStructure.RootList
            Dim work
            For Each work In System.IO.Directory.GetDirectories(root, "work")
                Dim proj
                For Each proj In System.IO.Directory.GetDirectories(work)
                    Dim sess
                    For Each sess In System.IO.Directory.GetDirectories(proj)
                        objFCSList = New FCS_List(sess)
                        If (System.IO.Directory.GetFiles(objFCSList.List_Spec).Length > 0) Then
                            If (Check_Work(System.IO.Directory.GetFiles(objFCSList.List_Spec)(0))) Then
                                mAnyWorkPath = sess
                                Any_Work = True
                            End If
                        End If
                    Next
                Next
            Next
        Next

    End Function

    ' dgp rev 3/29/2011 is a work merge reuired
    Public Shared Function IsPersonal(ByVal path As String) As Boolean

        If Not Directory.Exists(path) Then Return False
        Return path.ToLower.Contains("users") And path.ToLower.Contains(mUsername)

    End Function

    ' dgp rev 3/30/2011 Check Personal work space for last project and session
    Private Shared Function GetPersonal(ByVal path As String)

        Dim proj = System.IO.Path.GetFileName(Directory.GetParent(mLastWork.ToString).ToString)
        Dim sess = System.IO.Path.GetFileName(mLastWork.ToString)
        Dim arr = mLastWork.Split(mSep)
        Dim newarr(6) As String
        newarr(0) = arr(0)
        newarr(1) = arr(1)
        newarr(2) = "Users"
        newarr(3) = mUsername
        newarr(4) = "Work"
        newarr(5) = proj
        newarr(6) = sess
        Dim newwork = Join(newarr, mSep)
        If Directory.Exists(newwork) Then Return newwork
        Return Nothing

    End Function

    Private Shared mLastWork As String
    Private Shared mLastWorkValid As Boolean = False

    ' dgp rev 5/12/2011 one time establish (OTE) current work
    Private Shared Sub EstablishCurWork()

        mOTE_CurWork = True
        ' attribute is assigned to last work, then validated as personal
        If LastWorkExists() Then
            mCurWorkValid = True
            mCurWork = mLastWork
        End If

    End Sub

    Private Shared mOTE_CurWork As Boolean = False

    Public Shared Property CurWorkValid As Boolean
        Get
            If Not mOTE_CurWork Then EstablishCurWork()
            Return mCurWorkValid
        End Get
        Set(ByVal value As Boolean)

        End Set
    End Property


    ' dgp rev 7/16/08 Current Data Run
    Public Shared Property CurWork() As String
        Get
            If Not mOTE_CurWork Then EstablishCurWork()
            Return mCurWork
        End Get

        Set(ByVal value As String)
            mCurWork = value
            If (mWorkState = mState.None) Then
                ' dgp rev 7/16/08 no root, so no data
                mCurWorkValid = False
            Else
                mCurWorkValid = (System.IO.Directory.Exists(mCurWork))
                ' dgp rev 7/16/08  
                If (mCurDataValid) Then
                    mWorkState = mState.Populated
                    mNoWork = (Not mWorkState = mState.Populated)
                End If
            End If
        End Set
    End Property

    ' dgp rev 7/16/08 Current Data Run
    Public Shared Property CurSettings() As String
        Get
            Return mCurSettings
        End Get
        Set(ByVal value As String)
            mCurSettings = value
            If (mSettingsState = mState.None) Then
                ' dgp rev 7/16/08 no root, so no data
                mCurSettingsValid = False
            Else
                mCurSettingsValid = (System.IO.Directory.Exists(mCurSettings))
                ' dgp rev 7/16/08  
                If (mCurDataValid) Then
                    mSettingsState = mState.Populated
                    mNoSettings = (Not mSettingsState = mState.Populated)
                End If
            End If
        End Set
    End Property

    ' dgp rev 7/16/08 Current Data Run
    Public Shared Property CurData() As String
        Get
            Return mCurData
        End Get
        Set(ByVal value As String)
            mCurData = value
            mDataChange = True
            If (mDataState = mState.None) Then
                ' dgp rev 7/16/08 no root, so no data
                mCurDataValid = False
            Else
                mCurDataValid = (System.IO.Directory.Exists(System.IO.Path.Combine(mDataRoot, mCurData)))
                ' dgp rev 7/16/08  
                If (mCurDataValid) Then
                    mDataState = mState.Populated
                    mNoData = (Not mDistState = mState.Populated)
                End If
            End If
        End Set
    End Property

    Private Shared Function OfflineDistribution() As Boolean

        OfflineDistribution = False
        If System.IO.Directory.Exists(FlowStructure.FlowRoot) Then
            Dim path = System.IO.Path.Combine(FlowStructure.FlowRoot, "Distribution")
            If System.IO.Directory.Exists(path) Then
                Return System.IO.Directory.GetDirectories(path).Length > 0
            End If

        End If


    End Function

    Private Shared Function IsVersionLocal(ByVal version) As Boolean

        IsVersionLocal = False
        If System.IO.Directory.Exists(FlowStructure.FlowRoot) Then
            Dim path = System.IO.Path.Combine(FlowStructure.FlowRoot, "Distribution")
            If System.IO.Directory.Exists(path) Then
                Dim vers = System.IO.Path.Combine(path, version)
                Return System.IO.Directory.Exists(vers)
            End If

        End If

    End Function

    Private Shared Sub ConfigureVersion()

        mCurDistFlag = True

        If FlowStructure.FlowRoot Is Nothing Then Return

        If Not System.IO.Directory.Exists(System.IO.Path.Combine(FlowStructure.FlowRoot, "Distribution")) Then Return

        ' dgp rev 8/8/2011 if server is up, use current server version
        If DistributionServer.ServerUp Then
            If IsVersionLocal(DistributionServer.CurrentVersion) Then
                mCurDist = DistributionServer.CurrentVersion
            Else
                If DistributionServer.DownloadSelectDist(DistributionServer.CurrentVersion) Then
                    mCurDist = DistributionServer.CurrentVersion
                End If
            End If
        Else
            ' dgp rev 8/8/2011 if server is not up, use current local version
            If OfflineDistribution() Then
                mCurDist = PersistantDist()
            End If
        End If

    End Sub

    Private Shared mCurDist = Nothing
    Private Shared mCurDistFlag As Boolean = False
    Public Shared ReadOnly Property CurDistFullSpec As String
        Get
            Return System.IO.Path.Combine(Dist_Root, CurDist)
        End Get
    End Property

    ' dgp rev 7/16/08 Current Distribution
    Public Shared Property CurDist() As String
        Get
            If mCurDist Is Nothing Then
                If PVWaveDistribution.DistributionLocalExists Then
                    mCurDist = PVWaveDistribution.DistributionLocalName
                Else
                    ConfigureVersion()
                End If
            End If
            Return mCurDist
        End Get

        Set(ByVal value As String)
            If (System.IO.Path.GetDirectoryName(value).Length = 0) Then
                value = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(mCurDist), value)
            End If

            mCurDist = value
            mCurDistValid = (System.IO.Directory.Exists(value))
            ' dgp rev 7/16/08  
            If (mCurDistValid) Then
                mDistState = mState.Populated
            End If
        End Set
    End Property

    ' dgp rev 2/22/08 Create a new Distribution path, if none exists

    ' dgp rev 8/6/08 Create the Depot Path in current root
    Private Shared Function GCR_Create() As Boolean

        ' dgp rev 7/16/08 if server has a dist, then create folder
        If (Utility.Create_Tree(System.IO.Path.Combine(mFlowRoot, "Depot"))) Then
            Depot_Root = System.IO.Path.Combine(mFlowRoot, "Depot")
            Return True
        End If
        Return False

    End Function

    ' dgp rev 6/22/09 Get Any data at all
    Private Shared mAnyData = Nothing
    Public Shared ReadOnly Property AnyDataExists() As Boolean
        Get
            Return AnyData IsNot Nothing
        End Get
    End Property
    ' dgp rev 6/22/09 Some Data exists
    Private Shared ReadOnly Property AnyData() As String
        Get
            If (mAnyData Is Nothing) Then mAnyData = GetAnyData()
            Return mAnyData
        End Get
    End Property

    Private Shared mAnyWork = Nothing
    ' dgp rev 6/22/09 Establish Any work at all
    Public Shared ReadOnly Property AnyWork() As String
        Get
            If (mAnyWork Is Nothing) Then
                mAnyWork = GetAnyWork()
                mNoWork = mAnyWork Is Nothing
            End If
            Return mAnyWork
        End Get
    End Property

    Private Shared mTestWorkRoot
    Private mTestWorkSession

    ' dgp rev 11/25/09 Create work if non exists
    Private Shared Function CreateWork(ByVal path) As Boolean

        ' dgp rev 6/22/09 Nothing found, so create valid work from existing data

        CreateWork = False
        Dim test = AnyData

        Dim new_name = Unique_Name()

        Dim Session = System.IO.Path.Combine(System.IO.Path.Combine(path, new_name), new_name)

        If (Utility.Create_Tree(Session)) Then
            Create_Data_List(Session, test)
            mAnyWorkPath = Session
            mCurWork = Session
            CreateWork = True
        End If

    End Function

    ' dgp rev 12/2/09 Establish Work Root
    Public Shared Sub EstablishWorkRoot()

        mOneTimeWorkRootCheck = True
        mWorkRoots = New ArrayList

        Dim FoundWork = False

        mWorkRoot = Nothing
        ' dgp rev 11/24/09 look for persistant data settings
        If (PersistantWorkRoot()) Then
            If (AnyValidWork(mTestWorkRoot)) Then
                mWorkRoot = mTestWorkRoot
                Return
            End If
        Else
            ' dgp rev 11/24/09 look for data in FlowRoot
            If FlowRoot Is Nothing Then Return
            mTestWorkRoot = System.IO.Path.Combine(FlowRoot, "Work")
            If (System.IO.Directory.Exists(mTestWorkRoot)) Then
                If System.IO.Directory.GetDirectories(mTestWorkRoot).Length = 0 Then
                    ' dgp rev 11/24/09 Create Work from current work
                Else
                    mWorkRoot = mTestWorkRoot
                    Return
                End If
            End If
        End If

    End Sub

    ' dgp rev 2/22/08 Establish the entire FlowRoot Structure
    Public Shared Function Establish_Flow() As Boolean

        Establish_Flow = False

        If (FlowRoot Is Nothing) Then Exit Function
        Log_Info("Root Established")

        If (CurDist Is Nothing) Then Exit Function
        Log_Info("distribution Established")

        Establish_Settings()
        If (Settings Is Nothing) Then
            Log_Info("No Settings Established")
        Else
            Log_Info("Settings Established")
        End If

        EstablishDataRoot()
        If (Data_Root Is Nothing) Then
            Log_Info("No Data Established")
        Else
            Log_Info("Data Established")
        End If

        ' dgp rev 5/11/2011 Remove work establishment from flow establishment
        ' dgp rev 5/11/2011 some applications don't use work -- User Protocol, FCS Upload
        '        EstablishWorkRoot()
        '        If (WorkRootExists) Then
        '        Log_Info("Work found")
        '        Else
        '        Log_Info("No work")
        '        End If

        Establish_Flow = CurState

    End Function

    ' dgp rev 2/20/08 Establish DataRoot
    ' 1) Check XML file, then check that path is still valid
    ' 2) Check FlowRoot for Data area
    ' 3) Scan for Data area
    ' 4) Create new Data area

    ' dgp rev 8/6/08 Establish a Depot area for pre-processing data
    Public Shared Function Establish_Depot() As Boolean

        Establish_Depot = True

        ' dgp rev 7/10/08 check previous settings
        If (GCR_XML()) Then Exit Function

        If (GCR_FlowRoot()) Then Exit Function

        If (GCR_Scan()) Then Exit Function

        ' dgp rev 7/10/08 create the empty structure
        If (GCR_Create()) Then Exit Function

        Establish_Depot = False

        Log_Info("Failed to establish Depot Root")

    End Function

    Private Shared mUsersRoot = Nothing
    Public Shared ReadOnly Property UsersRoot() As String
        Get
            If mUsersRoot Is Nothing Then mUsersRoot = Establish_Users()
            Return mUsersRoot
        End Get
    End Property

    ' dgp rev 6/22/09 Establish Users
    Private Shared Function Establish_Users() As String

        ' read previous setup and validate
        If (mXML.Exists("UsersRoot")) Then
            If (System.IO.Directory.Exists(mXML.GetSetting("UsersRoot"))) Then
                Return mXML.GetSetting("UsersRoot")
            Else
                ' dgp rev 7/16/08 no Data at this point
                Log_Info("UsersRoot no longer valid")
            End If
        End If

        If FlowRoot Is Nothing Then Return Nothing

        Dim tmp = System.IO.Path.Combine(FlowRoot, "Users")
        If Utility.Create_Tree(tmp) Then
            mXML.PutSetting("UsersRoot", tmp)
            Return tmp
        Else
            Return Nothing
        End If

    End Function

    ' dgp rev 11/24/09 Get XML Data Path
    Private Shared Function PersistantDataRoot() As Boolean

        mDataState = mState.None
        ' dgp rev 2/20/08 keep this simple, if XML settings exists, then read it.
        If (mXML.NewFlag) Then Return False

        ' read previous setup and validate
        If (mXML.Exists("DataRoot")) Then
            mDataState = Check_Current(mXML.GetSetting("DataRoot"))
            If (System.IO.Directory.Exists(mXML.GetSetting("DataRoot"))) Then
                mTestDataRoot = mXML.GetSetting("DataRoot")
                ' dgp rev 7/16/08 Data root may be emtpy or populated at this point
                mNoData = (Not mDataState = mState.Populated)
                Return Not mNoData
            Else
                ' dgp rev 7/16/08 no Data at this point
                mNoData = True
                Log_Info("DataRoot no longer valid")
                Return False
            End If
        End If
        Return False

    End Function

    ' dgp rev 11/24/09 Get XML Data Path
    ' dgp rev 12/2/09 
    Private Shared Function PersistantUserRoot() As Boolean

        mUserState = mState.None
        ' dgp rev 2/20/08 keep this simple, if XML settings exists, then read it.
        If (mXML.NewFlag) Then Return False

        ' read previous setup and validate
        If (mXML.Exists("UserRoot")) Then
            mUserState = Check_Current(mXML.GetSetting("UserRoot"))
            If (System.IO.Directory.Exists(mXML.GetSetting("UserRoot"))) Then
                mTestUserRoot = mXML.GetSetting("UserRoot")
                ' dgp rev 7/16/08 User root may be emtpy or populated at this point
                mNoUser = (Not mUserState = mState.Populated)
                Return Not mNoUser
            Else
                ' dgp rev 7/16/08 no User at this point
                mNoUser = True
                Log_Info("UserRoot no longer valid")
                Return False
            End If
        End If
        Return False

    End Function

    ' dgp rev 12/2/09 Get XML Work Path
    Private Shared Function PersistantWorkRoot() As Boolean

        mWorkState = mState.None
        ' dgp rev 2/20/08 keep this simple, if XML settings exists, then read it.
        If (mXML.NewFlag) Then Return False

        ' read previous setup and validate
        If (mXML.Exists("WorkRoot")) Then
            mWorkState = Check_Current(mXML.GetSetting("WorkRoot"))
            If (System.IO.Directory.Exists(mXML.GetSetting("WorkRoot"))) Then
                mTestWorkRoot = mXML.GetSetting("WorkRoot")
                ' dgp rev 7/16/08 Work root may be emtpy or populated at this point
                mNoWork = (Not mWorkState = mState.Populated)
                Return Not mNoWork
            Else
                ' dgp rev 7/16/08 no Work at this point
                mNoWork = True
                Log_Info("WorkRoot no longer valid")
                Return False
            End If
        End If
        Return False

    End Function


    ' dgp rev 4/24/09 Establish Data
    Public Shared Function Establish_Data() As Boolean

        Establish_Data = True

        ' dgp rev 7/10/08 check previous settings
        RaiseEvent FSRemoteEvent("Reading XML...")
        If (GDR_XML()) Then Exit Function

        If (GDR_FlowRoot()) Then Exit Function

        If (GDR_Scan()) Then Exit Function

        ' dgp rev 7/10/08 create the empty structure

        Establish_Data = False

        Log_Info("Failed to establish Data Root")

    End Function

    Public Shared Sub UpdatePersistantDist(ByVal name)

        Dim tst = System.IO.Path.Combine(FlowStructure.Dist_Root, name)
        If System.IO.Directory.Exists(tst) Then
            mXML.PutSetting("CurDist", tst)
            mPersistantDist = tst
        End If

    End Sub

    Public Shared Function EstablishDistribution(ByVal path As String) As Boolean

        EstablishDistribution = False
        If System.IO.Directory.Exists(path) Then
            EstablishDistribution = True
            mCurDist = System.IO.Path.GetFileName(path)
            mDistRoot = System.IO.Path.GetDirectoryName(path)
        End If

    End Function

    ' dgp rev 2/7/2012 Scan for persistant distribution
    Public Shared Function ReadPersistantDist() As Boolean

        ReadPersistantDist = False
        If mXML.Exists("CurDist") Then
            mPersistantDist = mXML.GetSetting("CurDist")
            ReadPersistantDist = True
        End If

    End Function

    ' dgp rev 2/7/2012 Scan for persistant distribution
    Private Shared Sub ScanPersistantDist()

        mPersistantDist = "Current"
        If mXML.Exists("CurDist") Then
            If System.IO.Directory.Exists(mXML.GetSetting("CurDist")) Then
                mPersistantDist = mXML.GetSetting("CurDist")
                Exit Sub
            End If
        End If

        mPersistantDist = AnyLocalDist()
        If mPersistantDist IsNot Nothing Then mXML.PutSetting("CurDist", mPersistantDist)

    End Sub


    Private Shared mPersistantDist = Nothing
    Public Shared ReadOnly Property PersistantDist() As String
        Get
            If mPersistantDist Is Nothing Then ScanPersistantDist()
            Return mPersistantDist

        End Get
    End Property

    ' dgp rev 6/24/09 
    Private Shared mAnyLocalDistPath = Nothing
    Private Shared mAnyLocalDistFlag = Nothing
    Private Shared Function AnyLocalDistExits() As Boolean

        If mAnyLocalDistFlag Is Nothing Then
            Dim tmp
            Dim item

            ' dgp rev 6/22/09 find a populated distribution
            For Each item In RootList
                tmp = System.IO.Path.Combine(item, "Distribution")
                If System.IO.Directory.Exists(tmp) Then
                    mAnyLocalDistFlag = (System.IO.Directory.GetDirectories(tmp).Length > 0)
                    If (mAnyLocalDistFlag) Then
                        mXML.PutSetting("CurDist", System.IO.Directory.GetDirectories(tmp)(System.IO.Directory.GetDirectories(tmp).Length - 1))
                        mAnyLocalDistPath = System.IO.Directory.GetDirectories(tmp)(System.IO.Directory.GetDirectories(tmp).Length - 1)
                    End If
                End If
            Next
        End If
        Return mAnyLocalDistFlag

    End Function

    ' dgp rev 6/24/09 
    Private Shared Function AnyLocalDist() As String

        ' dgp rev 6/24/09 if any local distribution return, otherwise go to server
        If mAnyLocalDistPath Is Nothing Then
            If Not AnyLocalDistExits() Then
                mAnyLocalDistPath = DistributionServer.DownloadAnyServerDist()
            End If
        End If
        Return mAnyLocalDistPath

    End Function

    ' dgp rev 6/22/09 Scan for Valid Work in a given work root
    Private Shared Function ScanValidWork(ByVal workroot) As Boolean

        ScanValidWork = False
        mAnyWorkPath = Nothing
        Dim proj
        If Not System.IO.Directory.Exists(workroot) Then Return False
        If Not System.IO.Directory.GetDirectories(workroot).Length > 0 Then Return False
        ' dgp rev 9/29/2010 Use the FCS List object and not hard code
        Dim objFCSList As FCS_List

        For Each proj In System.IO.Directory.GetDirectories(workroot)
            Dim sess
            For Each sess In System.IO.Directory.GetDirectories(proj)
                objFCSList = New FCS_List(sess)
                If objFCSList.AnyValid Then
                    mAnyWorkPath = sess
                    ScanValidWork = True
                End If
            Next
        Next

    End Function

    ' dgp rev 1/23/2012 
    Private Shared mBetaInfo = Nothing
    Public Shared ReadOnly Property BetaInfo() As Hashtable
        Get
            If mBetaInfo Is Nothing Then
                Dim path = UsersRoot
                If (System.IO.Directory.Exists(path)) Then
                    path = System.IO.Path.Combine(path, mUsername + ".cfg")
                    If (System.IO.File.Exists(path)) Then
                        mBetaInfo = path
                    End If
                End If
            End If
            Return mBetaInfo
        End Get
    End Property

    ' dgp rev 6/23/09 Get Any Valid Work
    Private Shared Function GetAnyWork() As String

        If FlowRoot Is Nothing Then Return Nothing

        ' read previous setup and validate
        Log_Info("Get Any Work")
        If (mXML.Exists("WorkRoot")) Then
            ' dgp rev 6/22/09 Work root must have valid session
            If (ScanValidWork(mXML.GetSetting("WorkRoot"))) Then
                Return mAnyWorkPath
            Else
                ' dgp rev 7/16/08 no Work at this point
                Log_Info("WorkRoot no longer valid")
            End If
        End If

        Dim root
        For Each root In RootList
            Dim work
            For Each work In System.IO.Directory.GetDirectories(root, "work")
                If (Not work.ToLower = mXML.GetSetting("WorkRoot").ToLower) Then
                    If (ScanValidWork(work)) Then Return mAnyWorkPath
                End If
            Next
        Next

        ' dgp rev 6/22/09 Nothing found, so create valid work from existing data

        Dim test = AnyData

        Dim new_name = Unique_Name()

        Dim Session = System.IO.Path.Combine(System.IO.Path.Combine(Work_Root, new_name), new_name)

        If (Utility.Create_Tree(Session)) Then
            Create_Data_List(Session, test)
            mAnyWorkPath = Session
            Return Session
        End If

        Return Nothing

    End Function

    ' dgp rev 6/23/09 Get Any Valid Work
    Private Shared Function Establish_Settings() As String

        If FlowRoot Is Nothing Then Return Nothing

        ' read previous setup and validate
        If (mXML.Exists("SettingRoot")) Then
            ' dgp rev 6/22/09 Setting root must have valid session
            If (System.IO.Directory.Exists(mXML.GetSetting("SettingRoot"))) Then
                Return mXML.GetSetting("SettingRoot")
            End If
        End If

        ' dgp rev 6/29/09 
        If OldSettingsExists Then
            If PersonalSettingsExists Then
                ' dgp rev 6/26/09 both exist, use personal
                Return PersonalSettingsRoot
            Else
                Try
                    System.IO.Directory.Move(OldSettingsRoot, PersonalSettingsRoot)
                    Return PersonalSettingsRoot
                Catch ex As Exception
                    MsgBox("Error moving settings - " + ex.Message, MsgBoxStyle.Information)
                    If (Utility.Create_Tree(PersonalSettingsRoot)) Then
                        Return PersonalSettingsRoot
                    Else
                        MsgBox("Error creating settings path", MsgBoxStyle.Information)
                    End If
                End Try
            End If
        Else
            If PersonalSettingsExists Then
                Return PersonalSettingsRoot
            Else
                ' dgp rev 6/26/09 initial creation of directory
                If (Utility.Create_Tree(PersonalSettingsRoot)) Then
                    Return PersonalSettingsRoot
                Else
                    MsgBox("Error creating settings path", MsgBoxStyle.Information)
                End If
            End If
        End If
        Return Nothing

    End Function

    ' dgp rev 6/23/09 Get Any Valid Work
    Public Shared Function Existing_WorkRoot() As String

        If FlowRoot Is Nothing Then Return Nothing

        ' read previous setup and validate
        If (mXML.Exists("WorkRoot")) Then
            ' dgp rev 6/22/09 Setting root must have valid session
            If (System.IO.Directory.Exists(mXML.GetSetting("WorkRoot"))) Then
                Return mXML.GetSetting("WorkRoot")
            End If
        End If

        ' dgp rev 12/2/09 Find personal work
        If PersonalWorkRootExists Then
            ' dgp rev 6/26/09 both exist, use personal
            If (AnyValidWork(PersonalWorkRoot)) Then Return PersonalWorkRoot
        End If

        ' dgp rev 12/2/09 Move general work
        If System.IO.Directory.Exists(System.IO.Path.Combine(FlowRoot, "Work")) Then
            ' dgp rev 6/26/09 both exist, use personal
            If (AnyValidWork(System.IO.Path.Combine(FlowRoot, "Work"))) Then

                Utility.DirectoryCopy(System.IO.Path.Combine(FlowRoot, "Work"), PersonalWorkRoot)
                Return PersonalWorkRoot
            End If
        End If

        If (CreateWork(PersonalWorkRoot)) Then Return PersonalWorkRoot

        Return Nothing

    End Function



    ' dgp rev 6/23/09 Get Any Valid Work
    Private Shared Function Establish_WorkRoot() As String

        mWorkRoot_OTE = True
        If FlowRoot Is Nothing Then Return Nothing

        ' read previous setup and validate
        If (mXML.Exists("WorkRoot")) Then
            ' dgp rev 6/22/09 Setting root must have valid session
            If (System.IO.Directory.Exists(mXML.GetSetting("WorkRoot"))) Then
                Return mXML.GetSetting("WorkRoot")
            End If
        End If

        Dim arr = FlowRoot.Split(mSep)
        Dim newarr(4) As String
        newarr(0) = arr(0)
        newarr(1) = arr(1)
        newarr(2) = "Users"
        newarr(3) = mUsername
        newarr(4) = "Work"

        Return Join(newarr, mSep)

        If (CreateWork(PersonalWorkRoot)) Then Return PersonalWorkRoot


        ' dgp rev 12/2/09 Find personal work
        If PersonalWorkRootExists Then
            ' dgp rev 6/26/09 both exist, use personal
            If (AnyValidWork(PersonalWorkRoot)) Then Return PersonalWorkRoot
        End If


        Return Nothing

    End Function

    ' dgp rev 3/29/2011 Generic Work Root found
    Public Shared ReadOnly Property GenericWorkRoot As String
        Get
            If GenericWorkRootExists Then Return System.IO.Path.Combine(FlowRoot, "Work")
            Return ""
        End Get
    End Property

    ' dgp rev 3/29/2011 Generic Work Root found
    Public Shared ReadOnly Property GenericWorkRootExists As Boolean
        Get
            Return Directory.Exists(System.IO.Path.Combine(FlowRoot, "Work"))
        End Get
    End Property

    ' dgp rev 6/14/2010 Current Work State
    Public Shared ReadOnly Property WorkState As Int16
        Get
            Dim val = 0
            If (PersonalWorkRootExists) Then val = 1
            If (GenericWorkRootExists) Then val = val + 2
            Return val
        End Get
    End Property

    ' dgp rev 6/23/09 New Personal Work Format
    Private Shared mPersonalWorkRoot = Nothing
    Public Shared ReadOnly Property PersonalWorkRootExists() As Boolean
        Get
            Return System.IO.Directory.Exists(PersonalWorkRoot)
        End Get
    End Property

    Public Shared ReadOnly Property PersonalWorkRoot() As String
        Get

            If mPersonalWorkRoot Is Nothing Then mPersonalWorkRoot = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.Combine(FlowRoot, "Users"), mUsername), "Work")
            Return mPersonalWorkRoot
        End Get
    End Property

    Private Shared mMachineName As String = Environment.MachineName
    Public Shared ReadOnly Property MachineName As String
        Get
            Return mMachineName
        End Get
    End Property

    ' dgp rev 3/29/2011 Machine work root exists
    Public Shared ReadOnly Property WorkMachineRootExists As Boolean
        Get
            Return Directory.Exists(WorkMachineRoot)
        End Get
    End Property

    ' dgp rev 3/29/2011 Machine work root path
    Public Shared ReadOnly Property WorkMachineRoot As String
        Get
            Return System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.Combine(FlowRoot, "Users"), MachineName), "Work")
        End Get
    End Property

    ' dgp rev 3/29/2011 move all generic work projects
    Public Shared Function MoveGenericProjects()

        Dim projpath
        Dim projname
        For Each projpath In Directory.GetDirectories(GenericWorkRoot)
            projname = Path.GetFileName(projpath)
            If Directory.Exists(Path.Combine(PersonalWorkRoot, projname)) Then
                If Directory.Exists(Path.Combine(WorkMachineRoot, projname)) Then
                Else
                    Try
                        If (Utility.Create_Tree(WorkMachineRoot)) Then MoveGenericProjects = Movefiles(projpath, Path.Combine(WorkMachineRoot, projname))
                    Catch ex As Exception

                    End Try
                End If
            Else
                Try
                    If (Utility.Create_Tree(PersonalWorkRoot)) Then MoveGenericProjects = Movefiles(projpath, Path.Combine(PersonalWorkRoot, projname))
                Catch ex As Exception

                End Try
            End If
        Next
        Try
            Directory.Delete(GenericWorkRoot)
        Catch ex As Exception
            Return False
        End Try
        Return Not GenericWorkRootExists


    End Function

    Private Shared Function Movefiles(ByVal source, ByVal target) As Boolean

        If Directory.Exists(source) Then
            Try
                RaiseEvent TransferSource(source)
                RaiseEvent TransferTarget(target)
                System.IO.Directory.Move(source, target)
            Catch ex As Exception
                RaiseEvent TransferEvent(False)
                Return False
            End Try
        Else
            RaiseEvent TransferEvent(False)
            Return False
        End If
        RaiseEvent TransferEvent(True)
        Return True

    End Function

    ' dgp rev 6/23/09 Upgrade work to personal
    Public Shared Function UpgradeWork() As Boolean

        UpgradeWork = True

        If GenericWorkRootExists Then
            If PersonalWorkRootExists Then
                UpgradeWork = MoveGenericProjects()
                If WorkMachineRootExists Then
                    ' all full

                Else
                    Try
                        If (HelperClasses.Utility.Create_Tree(Directory.GetParent(WorkMachineRoot).ToString)) Then UpgradeWork = Movefiles(GenericWorkRoot, WorkMachineRoot)
                    Catch ex As Exception
                        UpgradeWork = False
                    End Try
                End If
            Else
                UpgradeWork = Movefiles(GenericWorkRoot, PersonalWorkRoot)
            End If
        End If

    End Function

    ' dgp rev 6/23/09 Old Settings Format
    Private Shared mOldSettingsRoot = Nothing
    Private Shared ReadOnly Property OldSettingsExists() As Boolean
        Get
            Return System.IO.Directory.Exists(OldSettingsRoot)
        End Get
    End Property
    Private Shared ReadOnly Property OldSettingsRoot() As String
        Get
            If mOldSettingsRoot Is Nothing Then mOldSettingsRoot = System.IO.Path.Combine(FlowRoot, "Settings")
            Return mOldSettingsRoot
        End Get
    End Property

    ' dgp rev 6/23/09 New Personal Settings Format
    Private Shared mPersonalSettingsRoot = Nothing
    Public Shared ReadOnly Property PersonalSettingsExists() As Boolean
        Get
            Return System.IO.Directory.Exists(PersonalSettingsRoot)
        End Get
    End Property
    Public Shared ReadOnly Property PersonalSettingsRoot() As String
        Get
            If mPersonalSettingsRoot Is Nothing Then
                Dim path = System.IO.Path.Combine(FlowRoot, "Users")
                path = System.IO.Path.Combine(path, mUsername)
                path = System.IO.Path.Combine(path, "Settings")
                mPersonalSettingsRoot = path
            End If
            Return mPersonalSettingsRoot
        End Get
    End Property

    ' dgp rev 6/19/09 Status of PVW XML file
    Private Shared mPVW_XML_exists = Nothing
    Public Shared ReadOnly Property PVW_XML_exists() As Boolean
        Get
            If mPVW_XML_exists Is Nothing Then mPVW_XML_exists = System.IO.File.Exists(PVW_XML_File)
            Return mPVW_XML_exists
        End Get
    End Property

    ' dgp rev 6/19/09 Status of PVW XML file
    Private Shared mPVW_CFG_exists = Nothing
    Public Shared ReadOnly Property PVW_CFG_exists() As Boolean
        Get
            Return System.IO.File.Exists(PVW_CFG_File)
        End Get
    End Property

    ' dgp rev 6/19/09 Status of PVW XML file
    Private Shared mPVW_XML_file As String
    Public Shared ReadOnly Property PVW_XML_File() As String
        Get
            If mPVW_XML_file Is Nothing Then
                Dim username As String = System.Environment.GetEnvironmentVariable("username")
                mPVW_XML_file = System.IO.Path.Combine(Utility.MyAppPath, username + ".xml")
            End If
            Return mPVW_XML_file
        End Get
    End Property

    ' dgp rev 6/19/09 Status of PVW XML file
    Private Shared mPVW_CFG_file As String
    Public Shared ReadOnly Property PVW_CFG_File() As String
        Get
            If mPVW_CFG_file Is Nothing Then
                If FlowRoot IsNot Nothing Then
                    Dim username As String = System.Environment.GetEnvironmentVariable("username")
                    mPVW_CFG_file = System.IO.Path.Combine(UsersRoot, username + ".cfg")
                End If
            End If
            Return mPVW_CFG_file
        End Get
    End Property

    ' dgp rev 6/19/09 PVW Load
    Private Shared mPVWSetup = Nothing
    Private Shared Function PVWLoad() As Hashtable

        PVWLoad = New Hashtable
        If PVW_XML_exists Then
            Log_Info("Reading XML " + PVW_XML_File)
            Dim PVW_XML_Doc = New XmlDocument
            Dim XMLNode As XmlNode
            Try
                PVW_XML_Doc.Load(PVW_XML_File)
            Catch ex As Exception
                Exit Function
            End Try
            XMLNode = PVW_XML_Doc.SelectSingleNode("Environment/FlowControl/Settings")
            If (Not XMLNode Is Nothing) Then
                Dim idx
                For idx = 0 To XMLNode.Attributes.Count - 1
                    PVWLoad.Add(XMLNode.Attributes.Item(idx).Name, XMLNode.Attributes.Item(idx).Value)
                    Log_Info(XMLNode.Attributes.Item(idx).Name + " => " + XMLNode.Attributes.Item(idx).Value)
                Next
            End If
        ElseIf PVW_CFG_exists Then
            ' dgp rev 11/1/06 read the last configuration file in text format
            ' dgp rev 11/3/06 add a settings hash
            Dim line As String
            Dim fig_arr() As String
            Log_Info("Checking for old config...")
            Log_Info(PVW_CFG_File.ToString)
            Dim sr As New System.IO.StreamReader(PVW_CFG_File)
            While (Not sr.EndOfStream)
                line = sr.ReadLine()
                Log_Info(line)
                fig_arr = Split(line, Mid(line, 1, 1))
                PVWLoad.Add(fig_arr(1), fig_arr(2))
                Log_Info(fig_arr(1) + " => " + fig_arr(2))
            End While
        End If

    End Function

    ' dgp rev 6/19/09 PVWSetup from either XML or older CFG file
    Public Shared Property PVWSetup() As Hashtable
        Get
            If mPVWSetup Is Nothing Then mPVWSetup = PVWLoad()
            Return mPVWSetup
        End Get
        Set(ByVal value As Hashtable)
            mPVWSetup = value
        End Set
    End Property

    ' dgp rev 2/20/08 Establish FlowRoot
    ' Sets FlowRoot variable and saves in XML file.  
    Public Shared Function Establish_Root() As Boolean

        Establish_Root = False

        ' dgp rev 6/22/09 keep this simple, Check for previous PVW setup first
        If (mXML.Exists("FlowRoot")) Then
            If (System.IO.Directory.Exists(mXML.GetSetting("FlowRoot"))) Then
                mNoRoot = False
                Establish_Root = True
                mFlowRoot = mXML.GetSetting("FlowRoot")
                Exit Function
            Else
                Log_Info("FlowRoot no longer valid")
            End If
        End If
        ' dgp rev 6/19/09 no XML, so new environment

        If (Init_Root()) Then
            Establish_Root = True
            mXML.PutSetting("FlowRoot", mFlowRoot)
            Log_Info("Successful One-Time Initialization of PC")
        Else
            Log_Info("Initialization Failed - " + mMessage)
        End If

    End Function

    ' dgp rev 7/14/08 Create the initial FlowRoot structure
    Public Shared Function Initial_Setup() As Boolean

        Initial_Setup = False

    End Function

End Class
