' Name:     FCS Run Class
' Author:   Donald G Plugge
' Date:     3/22/06 
' Purpose:  Class to track information concerning a given FCS run
Imports System.IO
Imports System.Xml
Imports System.Security.Cryptography
Imports HelperClasses

' DataSet (flowroot\data),  BrowseRun(anywhere)

Public Class FCSRun

    Public Shared objImp As RunAs_Impersonator

    ' use delegates for Repeat
    Delegate Sub delRepeat(ByVal Xname As String)
    Public Event evtRepeat As delRepeat

    ' use delegates for Repeat
    Delegate Sub RenameEventHandler(ByVal OldName As String, ByVal NewName As String)
    Public Event RenameEvent As RenameEventHandler

    ' dgp rev 4/20/09 First attempt to work with an instance pointer
    Private mProtocol As FCSTable
    Public Property Protocol() As FCSTable
        Get
            Return mProtocol
        End Get
        Set(ByVal value As FCSTable)
            mParMatchOrder = Nothing
            mProtocol = value
        End Set
    End Property

    ' dgp rev 4/20/09 extracts protocol information from all files in run
    ' places information in Protocol and returns T/F
    Public Function ExtractRunProtocol() As Boolean

        ' extract the table from
        ExtractRunProtocol = False

        Protocol = Nothing
        ' must be a valid run
        If (Valid_Run()) Then
            ' extract matrix from run
            '            If (ProKeysExists(0)) Then
            If (FCS_Files(0).Table_Flag) Then

                ' create protocol from first FCS file
                Protocol = New FCSTable(FCS_Files(0).ProtocolOrder)

                Dim idx
                Dim myRow As DataRow
                For idx = 0 To FCS_cnt - 1
                    myRow = Protocol.myTable.NewRow()
                    myRow.ItemArray = FCS_Files(idx).ProtocolValues.ToArray
                    Protocol.myTable.Rows.Add(myRow)
                Next
                ExtractRunProtocol = True
            End If
        End If

    End Function

    ' dgp rev 8/6/2010
    Private Function TableEntry(ByVal TubeNum, ByVal ParInx) As String

        Dim colidx = Protocol.myTable.Columns.IndexOf(Protocol.ProKeys(ParMatchOrder(ParInx)))
        Dim Val = Protocol.myTable.Rows.Item(TubeNum - 1)(colidx)
        Return Val

    End Function

    Private mCurProtocolRow As Integer
    Private mCurMergeFile As FCS_File

    Private Sub UpdateHeader(ByVal Key, ByVal Val)

        If mCurMergeFile.Header.ContainsKey(Key) Then
            mCurMergeFile.Header.Item(Key) = Val
        Else
            mCurMergeFile.Header.Add(Key, Val)
        End If

    End Sub

    ' dgp rev 12/9/2010 Merge the Current File with the current protocol row
    Private Function MergeFile() As Boolean

        Dim paridx
        Dim OrigKey
        MergeFile = False
        ' dgp rev 10/20/2010 Add protocol name as flag that protocol merged
        UpdateHeader("EIB", Protocol.ProtocolName)

        Dim ParToCol As New Hashtable
        Dim item As DictionaryEntry
        For Each item In CrossIndex
            ParToCol.Add(item.Value, item.Key)
        Next

        For paridx = 1 To ParCnt
            OrigKey = String.Format("$P{0}N", paridx)
            If (ParToCol.ContainsKey(paridx - 1)) Then
                UpdateHeader(String.Format("EIB{0}H", paridx), Protocol.Get_Keys(ParToCol(paridx - 1)))
                UpdateHeader(String.Format("EIB{0}C", paridx), Protocol.myTable.Rows.Item(mCurProtocolRow - 1)(ParToCol(paridx - 1)))
            Else
                UpdateHeader(String.Format("EIB{0}H", paridx), mCurMergeFile.Header(OrigKey))
                UpdateHeader(String.Format("EIB{0}C", paridx), mCurMergeFile.Header(OrigKey))
            End If
        Next

        ' dgp rev 10/5/2010 
        Dim Header As ArrayList = Protocol.Get_Keys
        Dim colidx As Integer
        Dim SeqNum As Integer = 1
        For colidx = 0 To Header.Count - 1
            If Not CrossIndex.ContainsKey(colidx) Then
                UpdateHeader(String.Format("EIB{0}K", SeqNum), Header(colidx))
                UpdateHeader(String.Format("EIB{0}V", SeqNum), Protocol.myTable.Rows.Item(mCurProtocolRow - 1)(colidx))
                SeqNum += 1
            End If
            UpdateHeader(Header(colidx), Protocol.myTable.Rows.Item(mCurProtocolRow - 1)(colidx))
        Next

    End Function

    ' dgp rev 4/24/09 Internal merge
    Private Function mMerge(ByVal force As Boolean) As Boolean

        mMerge = False
        If (Protocol Is Nothing) Then Return False

        If (Protocol.myTable.Rows.Count = 0) Then Return False

        Dim file As String

        ' dgp rev 12/8/2010 
        Dim Header As New ArrayList
        Dim HeaderName = ""
        For Each file In System.IO.Directory.GetFiles(mOrig_Path)
            mCurMergeFile = New FCS_File(file)
            If (mCurMergeFile.Valid) Then
                If (mCurMergeFile.TubeNumber(mCurProtocolRow)) Then
                    ' dgp rev 12/8/2010 for each row in the table
                    If mCurProtocolRow <= Protocol.myTable.Rows.Count Then
                        MergeFile()
                        mMerge = True
                    End If
                    mCurMergeFile.SwapByteFlag = Me.ByteSwapFlag
                    mCurMergeFile.Save_File(file)
                    mSavedFlag = True
                End If
            End If
        Next

    End Function

    Private mParKeys As Dictionary(Of String, String)

    ' dgp rev 8/6/2010
    Public ReadOnly Property ParKeys As Dictionary(Of String, String)
        Get
            Return mParKeys
        End Get
    End Property

    Private mParMatchLookup As Hashtable = Nothing
    Public Property ParMatchLookup As Hashtable
        Get
            Return mParMatchLookup
        End Get
        Set(ByVal value As Hashtable)
            mParMatchLookup = value
        End Set
    End Property

    Private mParMatchOrder As ArrayList = Nothing
    Public ReadOnly Property ParMatchOrder As ArrayList
        Get
            If mParMatchOrder Is Nothing Then ParameterMatches()
            Return mParMatchOrder
        End Get
    End Property


    Private mParList As ArrayList = Nothing
    Public ReadOnly Property ParList As ArrayList
        Get
            If mParList Is Nothing Then ParScan()
            Return mParList
        End Get
    End Property

    Private mParCnt As Int16
    Public ReadOnly Property ParCnt As Int16
        Get
            If mParCnt = -1 Then mParCnt = CInt(FirstFile.Header("$PAR"))
            Return mParCnt
        End Get
    End Property

    Private mFirstFile As FCS_File = Nothing
    Public ReadOnly Property FirstFile As FCS_File
        Get
            If mFirstFile Is Nothing Then GetFirst()
            Return mFirstFile
        End Get
    End Property

    Private Sub GetFirst()

        mFirstFile = New FCS_File(System.IO.Path.Combine(Orig_Path, FirstRunFile))

    End Sub

    ' dgp rev 8/6/2010 Scan Parameters for internal names
    Private Sub ParScan()

        If FirstFile.Header.Contains("$PAR") Then
            Dim idx
            Dim key
            mParList = New ArrayList
            For idx = 1 To ParCnt
                key = String.Format("$P{0}N", idx)
                If (FirstFile.Header.Contains(key)) Then
                    mParList.Add(FirstFile.Header(key))
                End If
            Next
        End If

    End Sub

    Private mMapper As ProtocolMapping

    Public Function ReMap(ByVal col, ByVal par) As Boolean

        Return mMapper.ReMap(col, par)

    End Function

    ' dgp rev 12/8/2010 Incorporate protocol with parameter names
    Public Sub IncorporateProtocol()

        mMapper = New ProtocolMapping

        ' get protocol header
        mMapper.Header(Protocol.Get_Keys())
        mMapper.ParNames(ParList)
        ' 
    End Sub

    Public ReadOnly Property CrossIndex As Hashtable
        Get
            Return mMapper.CrossIndex
        End Get
    End Property

    ' dgp rev 8/5/2010 
    Public Sub ParameterMatches()

        mParMatchOrder = New ArrayList
        mParMatchLookup = New Hashtable
        mParKeys = New Dictionary(Of String, String)

        If (mParList.Count = 0) Then Return

        Dim key
        ' dgp rev 8/5/2010 loop thru current protocol and look for matches
        Dim idx
        ' dgp rev 8/5/2010 loop thru parameter list and look for matches
        For idx = 0 To mParList.Count - 1
            key = FCSAntibodies.FindMatch(mParList(idx))
            If key <> "" And Not mParKeys.ContainsKey(key) Then
                mParKeys.Add(key, mParList(idx))
                mParMatchOrder.Add(key)
                mParMatchLookup.Add(key, idx)
            Else
                mParMatchOrder.Add("")
            End If
        Next

    End Sub

    Private mSavePath As String

    ' dgp rev 4/24/09 Cache the selected dataset
    ' dgp rev 9/30/09 Save the checksum path
    Public Function SaveInternal() As Boolean

        SaveInternal = False
        Dim objFCS As FCS_File
        Try
            Dim file
            For Each file In System.IO.Directory.GetFiles(mOrig_Path)
                objFCS = New FCS_File(file)
                If (objFCS.Valid) Then

                End If
            Next

            If mSavePath Is Nothing Then mSavePath = mOrig_Path
            If (FCS_cnt > 0) Then
                Dim spec
                For idx As Integer = 0 To FCS_cnt - 1
                    spec = System.IO.Path.Combine(mSavePath, Me.FCS_Files(idx).FileName)
                    FCS_Files(idx).SwapByteFlag = Me.ByteSwapFlag
                    FCS_Files(idx).Save_File(spec)
                    SaveInternal = True
                Next
            End If
        Catch ex As Exception
            SaveInternal = False
        End Try
        mSavedFlag = SaveInternal
        Return SaveInternal

    End Function

    Private mHeaderChangeFlag As Boolean = False
    Public ReadOnly Property HeaderChangedFlag As Boolean
        Get
            Return mHeaderChangeFlag
        End Get
    End Property

    Private mProtocolFlag As Boolean = False
    Public ReadOnly Property ProtocolFlag As Boolean
        Get
            Return (Protocol IsNot Nothing)
        End Get
    End Property

    Private mSavedFlag As Boolean = False

    Public ReadOnly Property MergedFlag As Boolean
        Get
            Return mHeaderChangeFlag And mSavedFlag
        End Get
    End Property

    ' dgp rev 4/24/09 Merge only if protocol matches run count
    Public Function Merge() As Boolean

        mHeaderChangeFlag = mMerge(True)
        Return mHeaderChangeFlag

    End Function

    ' dgp rev 4/20/09 Attempt to merge run with current protocol
    Public Function Merge(ByVal force As Boolean) As Boolean

        mHeaderChangeFlag = mMerge(True)
        Return mHeaderChangeFlag

    End Function

    ' dgp rev 5/22/09 
    Private mUpload_Info As FCSUpload
    Public Property Upload_Info() As FCS_Classes.FCSUpload
        Get
            Return mUpload_Info
        End Get
        Set(ByVal value As FCSUpload)
            mUpload_Info = value
        End Set
    End Property

    ' dgp rev 5/22/09 
    Public ReadOnly Property UnassignedName() As String
        Get
            Return Me.MDT_UR
        End Get
    End Property

    ' dgp rev 11/20/08 in order Assigned name or Basic machine date time name 
    Public ReadOnly Property RunName() As String
        Get
            If (Not Upload_Info Is Nothing) Then
                If (Not Upload_Info.RunSetFlag) Then
                    Return Upload_Info.AssignedRun
                End If
            End If
            Return UnassignedName
        End Get
    End Property

    ' dgp rev 7/12/07 Prefix state 
    Public Enum PrefixType
        [Date_ymd] = 0
        Date_mdy = 1
        NoDate = 2
    End Enum

    ' dgp rev 7/12/07 suffix state 
    Public Enum SuffixType
        [Sample] = 0
        FileOrder = 1
        BTIM = 2
        Original = 3
        ' dgp rev 2/7/2011 should only be in one location
        CellQuest = 4
    End Enum

    Public SuffixState As SuffixType = SuffixType.Sample
    Public PrefixState As PrefixType = PrefixType.Date_mdy

    Public RenameList As New Dictionary(Of String, String)

    ' dgp rev 4/21/09 Check for keys
    Public ReadOnly Property ProKeysExists(ByVal idx As Integer) As Boolean
        Get
            If (idx < 0 Or idx > FCS_cnt - 1) Then Return False
            Return (FCS_Files(idx).ProtocolExists)
        End Get
    End Property

    ' dgp rev 4/21/09 Protocol from first file
    Public ReadOnly Property ProKeys(ByVal idx As Integer) As ArrayList
        Get
            If (Not ProKeysExists(idx)) Then Return Nothing
            Return FCS_Files(idx).ProtocolKeys
        End Get
    End Property

    ' dgp rev 3/23/06 Extract keys from the file object
    Public Function ExtractProtocolKeys() As ArrayList

        If (ProKeysExists(0)) Then Return FCS_Files(0).ProtocolKeys

        Return Nothing

    End Function

    ' dgp rev 3/22/06 Extract the Table information from the FCS file headers
    Public Function ExtractFilesProtocol() As Boolean

        Protocol.Matrix = New ArrayList

        ' extract the run info into a matrix
        Dim fcs As FCS_File
        fcs = New FCS_File(FCS_List(0))
        If (fcs.ProtocolExists) Then
            Dim Values As New ArrayList
            Dim file
            For Each file In FCS_List
                fcs = New FCS_File(file.fullname)
                If (fcs.ProtocolExists) Then
                Else
                End If
            Next
            Return (Protocol.Matrix.Count > 0)
        End If
        Return False

    End Function

    Private mTableInfo As ArrayList = Nothing
    Private mAllKeys As ArrayList = Nothing

    Public ReadOnly Property TableInfo As ArrayList
        Get
            If mTableInfo Is Nothing Then ExtractKeys()
            Return mTableInfo
        End Get
    End Property

    Public ReadOnly Property AllKeys As ArrayList
        Get
            If mAllKeys Is Nothing Then ExtractKeys()
            Return mAllKeys
        End Get
    End Property

    ' dgp rev 8/2/2010 Retrieve the Keys
    Private Sub ExtractKeys()

        mTableInfo = New ArrayList
        mAllKeys = New ArrayList

        Dim item As String
        Dim found As Boolean = False
        Dim RowIndex = 1
        Dim FileIdx = 0
        Dim objFCS As FCS_File

        ' dgp rev 7/13/2010 how is the order of file in FCS_Files determined? 
        Dim file
        For Each file In System.IO.Directory.GetFiles(mOrig_Path)
            objFCS = New FCS_File(file)
            If (objFCS.Valid) Then
                For Each item In objFCS.Header.Keys
                    mAllKeys.Add(item)
                    If (item.StartsWith("#")) Then
                        mTableInfo.Add(item)
                        found = True
                    End If
                Next
                If found Then Exit For
            End If
        Next

    End Sub



    ' dgp rev 3/24/06 take information from the buffer matrix and
    ' load it into a DataTable.
    Public Function FCS_Matrix_to_Table() As Boolean

        Dim heading As New ArrayList
        Dim idx As Int16

        Protocol.myTable = New DataTable("FCS Table")

        ' create each column
        heading = Protocol.Matrix(0)
        For idx = 0 To heading.Count - 1
            Dim nc As New DataColumn
            If (heading.Item(idx) = "") Then heading.Item(idx) = "#Column" + CStr(idx + 1)
            nc.ColumnName = heading.Item(idx)
            nc.Caption = heading.Item(idx)
            Protocol.myTable.Columns.Add(nc)
        Next

        Dim file_idx, key_idx As Int16
        Dim myRow As DataRow

        ' fill in the rows
        For file_idx = 1 To Protocol.Matrix.Count - 1
            ' Once a table has been created, use the NewRow to create a DataRow.
            myRow = Protocol.myTable.NewRow()
            heading = Protocol.Matrix(file_idx)
            For key_idx = 0 To heading.Count - 1
                ' Then add the new row to the collection.
                myRow(key_idx) = heading(key_idx)
            Next
            Protocol.myTable.Rows.Add(myRow)
        Next
        Return ((Protocol.ParCnt * Protocol.Tubes) > 0)

    End Function

    Private mRealList As ArrayList
    Public ReadOnly Property RealList As ArrayList
        Get
            CreateRealList()
            Return mRealList
        End Get
    End Property

    ' dgp rev 6/3/09 Scan Data for current files
    Public Sub CreateRealList()

        mRealList = New ArrayList

        Dim item
        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            mRealList.Add(System.IO.Path.GetFileName(item))
        Next

    End Sub

    ' prefix using internal $DATE keyword
    Private Function Prefix() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim rdate As Date
        Dim DefDate As Date
        Dim sdate As String
        Dim frmt As String = "yyMMdd"
        If (PrefixState = PrefixType.Date_mdy) Then frmt = "MMddyy"

        Prefix = False

        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            If PrefixState = PrefixType.NoDate Then
                objFCS = New FCS_Classes.FCS_File(item)
                If (objFCS.Valid) Then
                    RenameList.Item(System.IO.Path.GetFileName(item)) = ""
                    Prefix = True
                End If
            Else
                DefDate = System.IO.File.GetCreationTime(item)
                objFCS = New FCS_Classes.FCS_File(item)
                If (objFCS.Valid) Then
                    If (objFCS.Header.Contains("$DATE")) Then
                        rdate = objFCS.Header("$DATE")
                    Else
                        rdate = DefDate
                    End If
                    sdate = Format(rdate, frmt)
                    RenameList.Item(System.IO.Path.GetFileName(item)) = sdate
                    Prefix = True
                End If
            End If

        Next

    End Function



    ' dgp rev 7/13/2010 Calculate the tube order
    Public Sub CalculateTubeOrder(ByVal path)

        mTubeIndex = New Collection
        Dim idx
        Dim val
        Dim objfcs As FCS_File
        For idx = 0 To Me.FCS_cnt - 1
            objfcs = FCS_Files(idx)
            If (objfcs.Valid) Then
                If (objfcs.Header.ContainsKey("TUBE NAME")) Then
                    val = objfcs.Header("TUBE NAME")
                    Dim arr = val.Split("_")
                    Dim num = 0
                    Dim isNum = Long.TryParse(arr(arr.length - 1), num)
                    If isNum Then
                        mTubeIndex.Add(idx, num.ToString)
                    End If
                End If
            End If
        Next

    End Sub


    ' dgp rev 7/13/2010 Calculate the tube order
    Public Sub CalculateTubeOrder()

        mTubeIndex = New Collection
        Dim idx
        Dim val
        Dim objfcs As FCS_File
        For idx = 0 To Me.FCS_cnt - 1
            objfcs = FCS_Files(idx)
            If (objfcs.Valid) Then
                If (objfcs.Header.ContainsKey("TUBE NAME")) Then
                    val = objfcs.Header("TUBE NAME")
                    Dim arr = val.Split("_")
                    Dim num = 0
                    Dim isNum = Long.TryParse(arr(arr.length - 1), num)
                    If isNum Then
                        mTubeIndex.Add(idx, num.ToString)
                    End If
                End If
            End If
        Next

    End Sub

    ' dgp rev 7/13/2010 tube order index into FCS files list
    Private mTubeIndex As Collection = Nothing
    Public ReadOnly Property TubeIndex As Collection
        Get
            If mTubeIndex Is Nothing Then CalculateTubeOrder()
            Return mTubeIndex
        End Get
    End Property

    ' dgp rev 7/13/2010 prefix using internal TUBE NAME or $SMNO keyword
    Private Function SeqTubeName() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim idx As Integer

        SeqTubeName = False

        idx = 0

        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                If (objFCS.Header.ContainsKey("TUBE NAME")) Then
                    val = objFCS.Header("TUBE NAME")
                    Dim arr = val.Split("_")
                    Dim num
                    Dim isNum = Long.TryParse(arr(arr.length - 1), num)
                    If isNum Then
                        RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + (String.Format("{0:D3}", CInt(num))) + ".fcs"
                        SeqTubeName = True
                    Else
                        RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + val + ".fcs"
                        SeqTubeName = True
                    End If
                ElseIf (objFCS.Header.ContainsKey("$SMNO")) Then
                    val = objFCS.Header("$SMNO")
                    RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + val
                    SeqTubeName = True
                Else

                End If
            End If
        Next

    End Function

    ' prefix using internal $FIL keyword
    Private Function SeqOrigName() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim idx As Integer

        SeqOrigName = False

        idx = 0

        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                idx = idx + 1
                If (objFCS.Header.ContainsKey("$FIL")) Then
                    val = objFCS.Header("$FIL")
                Else
                    val = (String.Format("{0:D3}", CInt(idx)))
                End If
                RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + val
                SeqOrigName = True
            End If
        Next

    End Function

    ' prefix using current directory order
    Private Function SeqFileOrder() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim count As Integer
        Dim seq As String

        SeqFileOrder = False
        count = 1

        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                seq = Format(Val(count), "000")
                count = count + 1
                RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + seq + ".fcs"
                SeqFileOrder = True
            End If
        Next

    End Function

    Private mBTimOrder As ArrayList
    Private mFirstBTim = Nothing

    Public ReadOnly Property FirstBTim() As String
        Get
            If Not mOneTimeScan Then ScanValidFCSFiles()
            If (mFirstBTim Is Nothing) Then BTimSort()
            Return mFirstBTim
        End Get
    End Property

    Private mDate As String
    ' sort by BTim
    Private Function BTimSort() As Boolean

        BTimSort = False

        ' dgp rev 5/26/09 Any Path?
        If (Not mPathExists) Then Return False

        ' dgp rev 5/26/09 Any Files?
        If (Not mFilesExist) Then Return False

        If m_fcsfiles Is Nothing Then Return False
        If m_fcsfiles.Count = 0 Then Return False
        If m_fcsfiles.Count = 1 Then Return True

        Dim item
        Dim objFCS As FCS_Classes.FCS_File

        Dim BTim As New Hashtable
        Dim SRT As New ArrayList

        Dim unsorted As New Hashtable
        Dim dup As New ArrayList

        Dim DateTracker As New ArrayList
        ' dgp rev 8/29/2011 Array for FCS_file(s)
        For Each objFCS In m_fcsfiles
            'objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                If (objFCS.Header.ContainsKey("$BTIM")) Then
                    If (unsorted.ContainsKey(objFCS.Header("$BTIM"))) Then
                        dup.Add(objFCS.FullSpec)
                    Else
                        unsorted.Add(objFCS.Header("$BTIM"), objFCS)
                        BTim.Add(objFCS.Header("$BTIM"), System.IO.Path.GetFileName(objFCS.FullSpec))
                        SRT.Add(objFCS.Header("$BTIM"))
                    End If
                End If
                If (objFCS.Header.ContainsKey("$DATE") And Not DateTracker.Contains(objFCS.Header("$DATE"))) _
                Then DateTracker.Add(objFCS.Header("$DATE"))
            End If
        Next

        SRT.Sort()
        mBTimOrder = New ArrayList
        Dim tmp = New ArrayList
        For Each item In SRT
            mBTimOrder.Add(BTim(item))
            tmp.Add(unsorted(item))
        Next

        m_fcsfiles = tmp
        unsorted = Nothing
        ' dgp rev 7/1/09 flag the run for multiple dates
        mMultipleDates = (DateTracker.Count > 1)
        BTimSort = (mBTimOrder.Count > 0)
        If (BTimSort) Then
            mFirstBTim = SRT(0)
            mDate = DateTracker.Item(0)
        End If

    End Function

    Private mMultipleDates = False

    ' prefix using internal $BTIM
    Private Function SeqBTim() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim idx As Integer

        Dim BTim As New Hashtable
        Dim SRT As New ArrayList

        SeqBTim = False

        idx = 0
        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                idx = idx + 1
                If (objFCS.Header.ContainsKey("$BTIM")) Then
                    BTim.Add(objFCS.Header("$BTIM"), System.IO.Path.GetFileName(item))
                    SRT.Add(objFCS.Header("$BTIM"))
                End If
            End If
        Next

        SRT.Sort()
        val = 1
        For Each item In SRT
            RenameList.Item(BTim(item)) = RenameList.Item(BTim(item)) + val + ".fcs"
            val = val + 1
        Next
        SeqBTim = True

    End Function

    ' dgp rev 5/22/09 
    Private mInnerRunName
    Public ReadOnly Property InnerRunName() As String
        Get
            If mInnerRunName Is Nothing Then
                If (Not Calc_Run_Name()) Then
                    Return ""
                End If
            End If
            Return mInnerRunName
        End Get
    End Property

    Private mHeaderRun As String
    ' dgp rev 5/20/09 Calculate the run name from header info
    Public Function Calc_Run_Name() As Boolean

        If (Not BTimSort()) Then Return False

        mMDT_UR = String.Format("{0}_{1}_{2}", SN, Format(Convert.ToDateTime(mDate), "MMM-dd-yy"), Format(Convert.ToDateTime(FirstBTim), "hh!mm"))

        Return True

    End Function

    Private mRenameList = Nothing

    Private Sub SortStandardNames()

        Dim BTim As New Hashtable
        mStandardSort = New ArrayList
        Dim SRT As New ArrayList
        Dim item As DictionaryEntry
        Dim objFCS As FCS_File
        For Each item In StandardNames
            Dim FullSpec As String = System.IO.Path.Combine(Orig_Path, item.Key)
            objFCS = New FCS_Classes.FCS_File(FullSpec)
            If (objFCS.Valid) Then
                If (objFCS.Header.ContainsKey("$BTIM")) Then
                    BTim.Add(objFCS.Header("$BTIM"), item.Key)
                    SRT.Add(objFCS.Header("$BTIM"))
                End If
            End If
        Next

        SRT.Sort()

        Dim name As String
        For Each name In SRT
            If BTim.ContainsKey(name) Then mStandardSort.Add(BTim(name))
        Next

    End Sub

    Private mStandardSort As ArrayList
    Public ReadOnly Property StandardSort As ArrayList
        Get
            SortStandardNames()
            Return mStandardSort
        End Get
    End Property

    Private mStandardNames As Hashtable = Nothing
    Public ReadOnly Property StandardNames As Hashtable
        Get
            'If mStandardNames Is Nothing Then CalcStandardNames()
            CalcStandardNames()
            Return mStandardNames
        End Get
    End Property

    ' prefix using internal $DATE keyword
    Private Function StandardPrefix() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim rdate As Date
        Dim DefDate As Date
        Dim sdate As String
        Dim frmt As String = "yyMMdd"
        If (PrefixState = PrefixType.Date_mdy) Then frmt = "MMddyy"

        StandardPrefix = False

        Dim FileName As String

        For Each item In System.IO.Directory.GetFiles(Orig_Path)
            FileName = System.IO.Path.GetFileName(item)
            If PrefixState = PrefixType.NoDate Then
                objFCS = New FCS_Classes.FCS_File(item)
                If (objFCS.Valid) Then
                    mStandardNames.Add(FileName, "")
                    StandardPrefix = True
                End If
            Else
                DefDate = System.IO.File.GetCreationTime(item)
                objFCS = New FCS_Classes.FCS_File(item)
                If (objFCS.Valid) Then
                    If (objFCS.Header.Contains("$DATE")) Then
                        rdate = objFCS.Header("$DATE")
                    Else
                        rdate = DefDate
                    End If
                    sdate = Format(rdate, frmt)
                    mStandardNames.Add(FileName, sdate)
                    StandardPrefix = True
                End If
            End If

        Next

    End Function

    ' dgp rev 7/13/2010 prefix using internal TUBE NAME or $SMNO keyword
    Private Function AddTubeName(ByVal NameHash As Hashtable) As Hashtable

        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim TmpHash As New Hashtable
        Dim FullSpec As String
        Dim item As DictionaryEntry
        Dim cnt = 0
        For Each item In NameHash
            FullSpec = System.IO.Path.Combine(Orig_Path, item.Key)
            objFCS = New FCS_Classes.FCS_File(FullSpec)
            If (objFCS.Valid) Then
                cnt += 1
                If (objFCS.Header.ContainsKey("TUBE NAME")) Then
                    val = objFCS.Header("TUBE NAME")
                    Dim arr = val.Split("_")
                    Dim num = 0
                    Dim isNum = Long.TryParse(arr(arr.length - 1), num)
                    If isNum Then
                        TmpHash.Add(item.Key, NameHash.Item(item.Key) + (String.Format("{0:D3}", CInt(num))) + ".fcs")
                    Else
                        TmpHash.Add(item.Key, NameHash.Item(item.Key) + val + ".fcs")
                    End If
                ElseIf (objFCS.Header.ContainsKey("$SMNO")) Then
                    val = objFCS.Header("$SMNO")
                    TmpHash.Add(item.Key, NameHash.Item(item.Key) + val + ".fcs")
                Else
                    If System.IO.Path.GetExtension(FullSpec).ToLower.Contains("fcs") Then
                        val = String.Format("{0:D3}.fcs", cnt)
                    Else
                        ' dgp rev 12/23/2010 assume extension contains number
                        val = System.IO.Path.GetExtension(FullSpec).Replace(".", "") + ".fcs"
                    End If
                    TmpHash.Add(item.Key, NameHash.Item(item.Key) + val)
                End If
            End If
        Next

        Return TmpHash

    End Function

    ' dgp rev 12/22/2010 
    Private Function AddBTim(ByVal Namehash As Hashtable) As Hashtable

        Dim item As DictionaryEntry
        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim idx As Integer

        Dim BTim As New Hashtable
        Dim SRT As New ArrayList
        Dim TmpHash As New Hashtable

        Dim FullSpec As String

        idx = 0
        For Each item In Namehash
            FullSpec = System.IO.Path.Combine(Orig_Path, item.Key)
            objFCS = New FCS_Classes.FCS_File(FullSpec)
            If (objFCS.Valid) Then
                idx = idx + 1
                If (objFCS.Header.ContainsKey("$BTIM")) Then
                    BTim.Add(objFCS.Header("$BTIM"), item.Key)
                    SRT.Add(objFCS.Header("$BTIM"))
                End If
            End If
        Next

        SRT.Sort()
        val = 1
        Dim Ordered As String
        For Each Ordered In SRT
            TmpHash.Add(BTim(Ordered), Namehash.Item(BTim(Ordered)) + (String.Format("{0:D3}", CInt(val))) + ".fcs")
            val = val + 1
        Next
        Return TmpHash

    End Function

    ' dgp rev 2/7/2011 base name on file extension
    Private Function AddExtName(ByVal Namehash As Hashtable) As Hashtable

        Dim objFCS As FCS_Classes.FCS_File
        Dim count As Integer = 1
        Dim seq As String
        Dim val As String = ""
        Dim idx As Integer = 0
        Dim TmpHash As New Hashtable
        Dim item As String
        Dim key As String
        Dim IsNum
        Dim Index = 0

        For Each item In System.IO.Directory.GetFiles(Orig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                Index += 1
                key = System.IO.Path.GetFileName(item)
                IsNum = Long.TryParse(System.IO.Path.GetExtension(item).Replace(".", ""), count)
                If Not IsNum Then count = Index
                seq = Format(count, "000")
                If Namehash.ContainsKey(key) Then TmpHash.Add(key, Namehash.Item(key) + seq + ".fcs")
            End If
        Next

        Return TmpHash

    End Function
    ' dgp rev 12/22/2010 
    Private Function AddFileOrder(ByVal Namehash As Hashtable) As Hashtable

        Dim objFCS As FCS_Classes.FCS_File
        Dim count As Integer = 1
        Dim seq As String
        Dim val As String = ""
        Dim idx As Integer = 0
        Dim TmpHash As New Hashtable
        Dim item As String
        Dim key As String

        For Each item In System.IO.Directory.GetFiles(Orig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                key = System.IO.Path.GetFileName(item)
                seq = Format(count, "000")
                count = count + 1
                If Namehash.ContainsKey(key) Then TmpHash.Add(key, Namehash.Item(key) + seq + ".fcs")
            End If
        Next

        Return TmpHash

    End Function

    ' dgp rev 12/22/2010 
    Private Function AddOrigName(ByVal Namehash As Hashtable) As Hashtable

        Dim objFCS As FCS_Classes.FCS_File
        Dim count As Integer = 1
        Dim val As String = ""
        Dim idx As Integer = 0
        Dim TmpHash As New Hashtable
        Dim item As DictionaryEntry
        Dim FullSpec As String

        For Each item In Namehash
            FullSpec = System.IO.Path.Combine(Orig_Path, item.Key)
            objFCS = New FCS_Classes.FCS_File(FullSpec)
            If (objFCS.Valid) Then
                idx = idx + 1
                If (objFCS.Header.ContainsKey("$FIL")) Then
                    val = objFCS.Header("$FIL")
                Else
                    val = (String.Format("{0:D3}", CInt(idx)))
                End If
                ' dgp rev 12/23/2010 No .FCS added to original name
                TmpHash.Add(item.Key, Namehash.Item(item.Key) + val)
            End If
        Next

        Return TmpHash

    End Function



    ' calculate the new name using selected prefix and sequence
    Public Function CalcStandardNames() As Boolean

        mStandardNames = New Hashtable
        If (StandardPrefix()) Then
            Select Case SuffixState
                Case SuffixType.Sample
                    mStandardNames = AddTubeName(mStandardNames)
                Case SuffixType.FileOrder
                    mStandardNames = AddFileOrder(mStandardNames)
                Case SuffixType.BTIM
                    mStandardNames = AddBTim(mStandardNames)
                Case SuffixType.Original
                    mStandardNames = AddOrigName(mStandardNames)
                Case SuffixType.CellQuest
                    mStandardNames = AddExtName(mStandardNames)
            End Select
        End If
        Return Not (mStandardNames.Count = 0)

    End Function



    ' calculate the new name using selected prefix and sequence
    Public Function Calc_New_Names() As Boolean

        RenameList.Clear()
        Calc_New_Names = False

        Calc_New_Names = Prefix()

        If (Calc_New_Names) Then
            Select Case SuffixState
                Case SuffixType.Sample
                    Calc_New_Names = SeqTubeName()
                Case SuffixType.FileOrder
                    Calc_New_Names = SeqFileOrder()
                Case SuffixType.BTIM
                    Calc_New_Names = SeqBTim()
                Case SuffixType.Original
                    Calc_New_Names = SeqOrigName()
            End Select
        End If

    End Function

    Private mRenameCount = Nothing
    Public ReadOnly Property RenameCount()
        Get
            If mRenameCount Is Nothing Then Return 0
            Return mRenameCount
        End Get
    End Property

    Private mRenameErrors = Nothing
    Public ReadOnly Property RenameErrors()
        Get
            If mRenameErrors Is Nothing Then Return 0
            Return mRenameErrors
        End Get
    End Property

    'dgp rev 7/12/07 Rename the Origial Files
    Public Function StandardRename() As Boolean

        StandardRename = True

        Dim fil
        Dim new_name As String
        Dim item
        Dim prev As Integer = StandardNames.Count + 1

        Dim Remove_List As New ArrayList
        Dim Proc_List As Hashtable = StandardNames

        ' loop thru the list and remove successful renames
        ' dgp rev 7/11/07 loop inserted to handle repetitive rename for filealreadyexists
        mRenameCount = 0
        mRenameErrors = 0

        While (Remove_List.Count <> Proc_List.Count)
            Remove_List = New ArrayList
            If (Proc_List.Count = prev) Then Exit While
            prev = Proc_List.Count
            For Each item In Proc_List
                fil = System.IO.Path.Combine(Orig_Path, item.Key)
                new_name = System.IO.Path.Combine(Orig_Path, item.Value)
                Try
                    System.IO.File.Move(fil, new_name)
                    RaiseEvent RenameEvent(fil, new_name)
                    mRenameCount = mRenameCount + 1
                    Remove_List.Add(item.Key)
                Catch ex As Exception
                    '                    If (ex.Message.ToLower.Contains("filealreadyexists")) Then flgSpecCase = True
                    StandardRename = False
                    mRenameErrors = mRenameErrors + 1
                End Try
            Next
            Dim rem_item
            For Each rem_item In Remove_List
                If Proc_List.ContainsKey(rem_item) Then Proc_List.Remove(rem_item)
            Next
        End While
        ' dgp rev 7/23/2010 reset fcsfiles and fcslist so they are recalculated
        StandardRename = Proc_List.Count = 0

    End Function

    'dgp rev 7/12/07 Rename the Origial Files
    Public Function Rename_Files() As Boolean

        Rename_Files = True

        Dim fil
        Dim new_name As String
        Dim item
        Dim prev As Integer = RenameList.Count + 1

        Dim Remove_List As New ArrayList

        Dim Proc_List As Dictionary(Of String, String) = RenameList

        ' loop thru the list and remove successful renames
        ' dgp rev 7/11/07 loop inserted to handle repetitive rename for filealreadyexists
        '            For Each item In System.IO.Directory.GetFiles(RunCache.CacheDataPath)
        mRenameCount = 0
        mRenameErrors = 0

        While (Remove_List.Count <> Proc_List.Count)
            Remove_List = New ArrayList
            If (Proc_List.Count = prev) Then Exit While
            prev = Proc_List.Count
            For Each item In Proc_List
                fil = System.IO.Path.Combine(mOrig_Path, item.Key)
                new_name = System.IO.Path.Combine(mOrig_Path, item.Value)
                Try
                    System.IO.File.Move(fil, new_name)
                    RaiseEvent RenameEvent(fil, new_name)
                    mRenameCount = mRenameCount + 1
                    Remove_List.Add(item.Key)
                Catch ex As Exception
                    '                    If (ex.Message.ToLower.Contains("filealreadyexists")) Then flgSpecCase = True
                    Rename_Files = False
                    mRenameErrors = mRenameErrors + 1
                End Try
            Next
            Dim rem_item
            For Each rem_item In Remove_List
                Proc_List.Remove(rem_item)
            Next
        End While
        ' dgp rev 7/23/2010 reset fcsfiles and fcslist so they are recalculated
        If (Proc_List.Count = 0) Then
            '                m_fcsfiles = Nothing
            '                mDirInfo = Nothing
            '                m_fcslist = Nothing
            '                mInitRun
            Rename_Files = True
        End If

    End Function

    ' dgp rev 7/30/08 NCI User name
    Private mNCIUser = Nothing

    ' dgp rev 6/3/09 Clear the Current user
    Public Sub ClearUser()

        mNCIUser = Nothing

    End Sub


    Private mUserAssigned = False
    Public ReadOnly Property UserAssigned() As Boolean
        Get
            Return mUserAssigned
        End Get
    End Property

    ' dgp rev 6/3/09
    ' dgp rev 10/6/09 
    Public Property NCIUser() As String
        Get
            Return mNCIUser
        End Get
        Set(ByVal value As String)
            mNCIUser = value
        End Set
    End Property


    Public Run_List_Path As String = "\\Nt-eib-10-6b16\FTP_root\runs"

    ' dgp rev 7/29/08 Valid User consists of Valid Experiment and User from List
    Public Function Valid_User(ByVal test As String) As Boolean

        If (FCS_Classes.NIHNet.FlowLabUsers Is Nothing) Then Return False
        Return FCS_Classes.NIHNet.FlowLabUsers.Contains(test)

    End Function

    ' dgp rev 7/29/08 Valid User consists of Valid Experiment and User from List
    Public Function Valid_User() As Boolean

        If mNCIUser Is Nothing Then Return False
        If (FCS_Classes.NIHNet.FlowLabUsers Is Nothing) Then Return False
        Return FCS_Classes.NIHNet.FlowLabUsers.Contains(NCIUser)

    End Function

    ' dgp rev 7/11/07 upload data to server
    ' may need to use authentication
    ' dgp rev 7/27/07 Upload an FCS Run to Server
    ' dgp rev 3/4/09 Upload looks for the RESERVED run to replace
    ' dgp rev 3/5/09 Look for empty reserved folder and replace with data folder
    ' dgp rev 3/5/09 separate the building of the source and target from the transfer
    ' dgp rev rev 3/10/09 move upload routine to FCSUpload class
    ' dgp rev 8/6/08 
    Private Function CheckServer() As Boolean

        Dim Server_Path As String
        CheckServer = False

        If (Not Upload_Info Is Nothing) Then

            Server_Path = String.Format("\\{0}{1}", Upload_Info.DataServer, FCSUpload.Upload_Root)
            Server_Path = System.IO.Path.Combine(Server_Path, NCIUser)

            Server_Path = System.IO.Path.Combine(Server_Path, Unique_Prefix)
            FCSUpload.objImp.ImpersonateStart()

            CheckServer = System.IO.Directory.Exists(Server_Path)
            FCSUpload.objImp.ImpersonateStop()

        End If

    End Function

    Private mStoreRoot As String
    Public Property StoreRoot() As String
        Get
            Return mStoreRoot
        End Get
        Set(ByVal value As String)
            mStoreRoot = value
        End Set
    End Property

    Private mStoredFlag As Boolean = False

    Private mStatus As String

    ' dgp rev 8/6/08 Original Run Path
    Private mOrig_Path As String
    Private mValidRun = Nothing



    Private mValidPath As Boolean = False
    Private mDirInfo As DirectoryInfo
    Private mDataHomeFlag As Boolean = False

    Private mPathExists As Boolean = False
    Private mFilesExist As Boolean = False

    ' dgp rev 5/19/09 
    Private Sub mInitRun(ByVal pathspec As String)

        mUpload_Info = Nothing
        mChkSumHash = Nothing

        mParCnt = -1

        mOrig_Path = pathspec
        ' dgp rev 5/19/09 at a minimum, path must exist and contain files
        mPathExists = System.IO.Directory.Exists(pathspec)
        If (Not mPathExists) Then Return

        mFilesExist = (Not System.IO.Directory.GetFiles(pathspec).Length = 0)
        If Not mFilesExist Then Return

        mDirInfo = New DirectoryInfo(pathspec)
        If (mDirInfo.Parent.Name.ToLower = "data") Then If (mDirInfo.Parent.Parent.Name.ToLower = "flowroot") Then mDataHomeFlag = True

        m_FCS_cnt = Nothing
        mValidRun = Nothing
        mValidPath = True

    End Sub

    Public ReadOnly Property Orig_Path() As String
        Get
            Return mOrig_Path
        End Get
    End Property
    ' dgp rev 6/3/09 Users of this data
    Private mUsers = Nothing
    Public ReadOnly Property Users() As ArrayList
        Get
            If mUsers Is Nothing Then mUsers = FindUsers()
            mUsers = FindUsers()
            Return mUsers
        End Get
    End Property


    ' dgp rev 1/13/09 VMS User
    Private mMachine As String
    Public mDynRun As Dynamic

    Public Property DataInfo() As Dynamic
        Get
            Return mDynRun
        End Get
        Set(ByVal value As Dynamic)
            mDynRun = value
        End Set
    End Property

    ' dgp rev 5/27/09 Return Status
    Public ReadOnly Property Status() As String
        Get
            Return mStatus
        End Get
    End Property

    ' dgp rev 6/5/09 Final Run Name on Server
    Private mFinalRunName = Nothing
    Public ReadOnly Property FinalRunName() As Run_Name
        Get
            Return mFinalRunName
        End Get
    End Property

    ' dgp rev 6/4/09 Uploaded Run is found on Server
    Private Function Upload_Exists(ByVal user As String) As Boolean

        Upload_Exists = False

        If (Upload_Info Is Nothing) Then Exit Function
        Dim Server_Path As String

        Server_Path = String.Format("\\{0}{1}", Upload_Info.DataServer, FCSUpload.Upload_Root)
        Server_Path = System.IO.Path.Combine(Server_Path, user)

        If (System.IO.Directory.Exists(Server_Path)) Then
            Dim di As New DirectoryInfo(Server_Path)
            Dim lst = di.GetDirectories(Unique_Prefix() + "*")
            If (lst.Length > 0) Then
                Upload_Exists = True
                mFinalRunName = New Run_Name(lst(0).Name)
                If (FinalRunName.Assigned) Then
                    Upload_Info.AssignedRun = FinalRunName.RunName.ToUpper
                    If (Not FinalRunName.User.ToLower = Upload_Info.AssignedMap.ToLower) Then
                        Upload_Info.AssignedMap = FinalRunName.User.ToUpper
                    End If
                    Upload_Info.AssignedUser = NCIUser.ToLower
                End If
            End If
        End If

    End Function

    ' dgp rev 6/2/09 Loop through and move only FCS files to new location
    Private Function Move_FCS(ByVal source, ByVal dest) As Boolean

        Dim objfcs As FCS_File
        Dim item
        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            objfcs = New FCS_Classes.FCS_File(item)
            If (objfcs.Valid) Then
                System.IO.File.Copy(item, System.IO.Path.Combine(dest, objfcs.FileName), True)
            End If
        Next

    End Function

    ' dgp rev 8/2/07 Check to see if key is a valid format for run (R00000)
    Private Function Is_Run(ByVal key As String) As Boolean

        Is_Run = False
        If (key.ToLower.Substring(0, 1) = "r") Then
            If (key.Length = 6) Then
                Try
                    Dim val As Integer = CInt(key.Substring(1, 5))
                    Is_Run = True
                Catch ex As Exception

                End Try
            End If
        End If

    End Function

    ' dgp rev 8/2/07 check to see if Aria data with Run and User
    Public Function IsAria() As Boolean

        Dim key As String
        Dim val As String
        Dim idx As Integer = 0
        Dim cnt As Integer
        key = "$SRC"

        IsAria = False

        If (FCS_Files(idx) Is Nothing) Then Exit Function
        ' if first file doesn't contain key, then none contain key
        If (Not FCS_Files(idx).Header.Contains(key)) Then Exit Function

        For idx = 0 To Me.FCS_cnt - 1
            val = FCS_Files(idx).Header(key)
            FCS_Files(idx).Header = Nothing
            If (Is_Run(val)) Then
                m_run_num = val
                Run_Flag = True
                cnt = idx
                IsAria = True
                Exit For
            End If
        Next

        ' if valid run, then retrieve username also
        If (Run_Flag) Then
            key = "EXPERIMENT NAME"
            If (FCS_Files(cnt).Header.Contains(key)) Then
                m_run_user = FCS_Files(cnt).Header(key)
            Else
                m_run_user = "UNK"
            End If
        End If

    End Function

    Public ReadOnly Property SN() As String
        Get
            Dim key As String
            m_run_machine = "Cytometer"
            If (Not FCS_Files(0) Is Nothing) Then
                key = "$CYT"
                If (Not FCS_Files(0).Header(key) = "") Then m_run_machine = FCS_Files(0).Header(key)
                key = "CYTNUM"
                If (Not FCS_Files(0).Header(key) = "") Then m_run_machine = FCS_Files(0).Header(key)
                key = "CYTSN"
                If (Not FCS_Files(0).Header(key) = "") Then m_run_machine = FCS_Files(0).Header(key)
            End If
            Return m_run_machine
        End Get
    End Property

    ' If the run has valid data, then extract the Cytometry Serial Number $CYTSN
    Private Sub Check_SN()
        Dim key As String
        If (Not FCS_Files(0) Is Nothing) Then
            key = "$CYT"
            If (Not FCS_Files(0).Header(key) = "") Then m_run_machine = FCS_Files(0).Header(key)
            key = "CYTNUM"
            If (Not FCS_Files(0).Header(key) = "") Then m_run_machine = FCS_Files(0).Header(key)
            key = "CYTSN"
            If (Not FCS_Files(0).Header(key) = "") Then m_run_machine = FCS_Files(0).Header(key)
            If (m_run_machine = "") Then m_run_machine = "Cytometer"
        End If

    End Sub
    ' If the run has valid data, then extract the collection date -- $DATE
    Private Sub Check_Date()
        Dim key As String
        If (Not FCS_Files(0) Is Nothing) Then
            key = "$DATE"
            If (Not FCS_Files(0).Header(key) = "") Then
                m_run_date = FCS_Files(0).Header(key)
                VMS_Name = Format(Convert.ToDateTime(m_run_date), "MMddyy")

                Console.WriteLine("Date is " + VMS_Name.ToString)
            End If
            If (m_run_date = "") Then m_run_date = "Date"
        End If

    End Sub
    ' If the run has valid data, then extract the collection time -- $BTIM
    Private Sub Check_Time()
        Dim key As String
        Dim Raw_Str As String()
        If (Not FCS_Files(0) Is Nothing) Then
            key = "$BTIM"
            If (Not FCS_Files(0).Header(key) = "") Then
                Raw_Str = FCS_Files(0).Header(key).Split(":")
                m_run_time = Raw_Str(0) + "!" + Raw_Str(1)
            End If
            If (m_run_time = "") Then m_run_time = "Time"
        End If

    End Sub

    ' dgp rev 6/11/09 list of checksums for all valid FCS files in the run
    Private mChkSumHash As New Hashtable
    Public ReadOnly Property ChkSum_List() As Hashtable
        Get
            If mChkSumHash Is Nothing Then Calc_ChkSum_List()
            Return mChkSumHash
        End Get
    End Property

    ' VMS File naming format is the Data as mmddyyxxx.fcs
    ' where xxx is the sequence number of the file
    ' dgp rev 5/18/06 
    Private m_vms_name As String
    Public Property VMS_Name() As String
        Get
            Return m_vms_name
        End Get
        Set(ByVal Value As String)
            m_vms_name = Value
        End Set
    End Property

    Private m_cur_index As Integer = 1
    Public Property Cur_Index() As String
        Get
            Return m_cur_index
        End Get
        Set(ByVal Value As String)
            m_cur_index = Value
        End Set
    End Property

    ' Run name is combination of machine, date, time, user and sequence number
    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_machine As String
    Public Property Run_Machine() As String
        Get
            If (m_run_machine Is Nothing) Then Check_SN()
            Return m_run_machine
        End Get
        Set(ByVal Value As String)
            m_run_machine = Value
        End Set
    End Property

    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_date As String
    Public Property Run_Date() As String
        Get
            If (m_run_date Is Nothing) Then Check_Date()
            Return m_run_date
        End Get
        Set(ByVal Value As String)
            m_run_date = Value
        End Set
    End Property

    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_time As String
    Public Property Run_Time() As String
        Get
            If (m_run_time Is Nothing) Then Check_Time()
            Return m_run_time
        End Get
        Set(ByVal Value As String)
            m_run_time = Value
        End Set
    End Property
    ' Flag if User and Run Number are to be used
    Private m_run_flag As Boolean = False
    Public Property Run_Flag() As Boolean
        Get
            Return m_run_flag
        End Get
        Set(ByVal value As Boolean)
            m_run_flag = value
        End Set
    End Property

    ' dgp rev 6/4/09 Set the User and Runname for a given dataset
    Public Sub Set_User_Run(ByVal user As String, ByVal run As String)

        Run_Flag = True
        Run_User = user
        Run_Num = run

    End Sub

    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_user As String
    Public Property Run_User() As String
        Get
            Return m_run_user
        End Get
        Set(ByVal Value As String)
            m_run_user = Value
        End Set
    End Property

    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_num As String
    Public Property Run_Num() As String
        Get
            Return m_run_num
        End Get
        Set(ByVal Value As String)
            m_run_num = Value
        End Set
    End Property

    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_vers As String
    Public Property Run_Vers() As String
        Get
            Return m_run_vers
        End Get
        Set(ByVal Value As String)
            m_run_vers = Value
        End Set
    End Property

    ' dgp rev 4/24/06 run - machine used to collect data
    Private m_run_id As String
    Public Property Run_Id() As String
        Get
            Return m_run_id
        End Get
        Set(ByVal Value As String)
            m_run_id = Value
        End Set
    End Property

    ' path of data file
    Private m_data_path As String
    Public Property Data_Path() As String
        Get
            If m_data_path = "" Then m_data_path = mOrig_Path
            Return m_data_path
        End Get
        Set(ByVal Value As String)
            m_data_path = Value
        End Set
    End Property

    Private m_fcs_array As New ArrayList
    ' number of files in the run
    Private m_FCS_cnt
    Public Property FCS_cnt() As Int16
        Get
            If m_FCS_cnt Is Nothing Then
                If Me.mFilesExist Then
                    m_FCS_cnt = FCS_List.Length
                Else
                    m_FCS_cnt = 0
                End If
            End If
            Return m_FCS_cnt
        End Get
        Set(ByVal Value As Int16)
            m_FCS_cnt = Value
        End Set
    End Property
    ' list of FCS file names
    Private m_fcslist As Array = Nothing
    Public Property FCS_List() As Array
        Get
            If m_fcslist Is Nothing Then m_fcslist = mDirInfo.GetFiles
            Return m_fcslist
        End Get
        Set(ByVal Value As Array)
            m_fcslist = Value
        End Set
    End Property

    Private mOneTimeScan As Boolean = False
    ' sort by BTim
    Private Sub ScanValidFCSFiles()

        mOneTimeScan = True
        Dim item
        Dim objFCS As FCS_Classes.FCS_File

        m_fcsfiles = New ArrayList

        If System.IO.Directory.Exists(mOrig_Path) Then
            For Each item In System.IO.Directory.GetFiles(mOrig_Path)
                objFCS = New FCS_Classes.FCS_File(item)
                If (objFCS.Valid) Then
                    m_fcsfiles.Add(objFCS)
                End If
            Next
        End If

    End Sub

    ' dgp rev 8/15/2011
    Public ReadOnly Property FileCount
        Get
            If m_fcsfiles Is Nothing Then Return 0
            Return m_fcsfiles.Count
        End Get
    End Property

    ' array of FCS files objects
    Private m_fcsfiles As ArrayList = Nothing
    Public Property FCS_Files(ByVal idx As Int16) As FCS_File
        Get
            If Not mOneTimeScan Then ScanValidFCSFiles()
            If idx > m_fcsfiles.Count - 1 Then Return Nothing
            Return m_fcsfiles(idx)
        End Get
        Set(ByVal Value As FCS_File)
            m_fcsfiles.Add(Value)
        End Set
    End Property

    Private mChkSumList As ArrayList

    ' dgp rev 5/9/06 Calculate Checksum
    Public Sub Calc_ChkSum_List()

        Dim objFCS As FCS_File
        Dim item As FileInfo
        Dim chksum As Byte()
        Dim bite As Byte
        Dim HEXStr As String

        mChkSumHash = New Hashtable
        mChkSumList = New ArrayList

        For Each item In FCS_List
            ' dgp rev 6/11/09 why .fcs extension, rather use Valid flag
            '            If (LCase(item.Extension) = ".fcs") Then
            HEXStr = ""
            objFCS = New FCS_File(item.FullName.ToString)
            If (objFCS.Valid) Then
                If objFCS.ValidData Then
                    chksum = objFCS.Calc_ChkSum
                    For Each bite In chksum
                        HEXStr = HEXStr + (String.Format("{0:X2}", bite))
                    Next
                    mChkSumList.Add(HEXStr.ToString)
                    If Not mChkSumHash.ContainsKey(HEXStr.ToString) Then mChkSumHash.Add(HEXStr.ToString(), item.Name)
                End If
            End If
            '            End If
        Next

    End Sub

    ' dgp rev 5/9/06 Calculate Checksum
    Public Function fn_ChkSum() As String

        If (FCS_Files(Cur_Index).Valid) Then

            Dim objFCS As FCS_File = FCS_Files(Cur_Index)
            Dim chksum As Byte() = objFCS.Calc_ChkSum
            Dim HEXStr As String = ""
            Dim bite As Byte

            For Each bite In chksum
                HEXStr = HEXStr + (String.Format("{0:X1}", bite))
            Next
            fn_ChkSum = HEXStr.ToString
        Else
            fn_ChkSum = "0"
        End If

    End Function

    ' dgp rev 5/16/06 unique checksum for each dataset
    Private m_chksum As String
    Public Property ChkSum() As String
        Get
            If (m_chksum Is Nothing) Then m_chksum = fn_ChkSum()
            Return m_chksum
        End Get
        Set(ByVal Value As String)
            m_chksum = Value
        End Set
    End Property

    ' dgp rev 5/17/07 Validate the run if at least one valid FCS file
    Public ReadOnly Property Valid_Run() As Boolean
        Get
            If (mValidRun Is Nothing) Then
                ' are their any files in the path?
                If (Me.mPathExists And Me.mFilesExist) Then
                    ' are any of the files proper FCS?
                    Dim item
                    Dim tmp As FCS_File
                    For Each item In System.IO.Directory.GetFiles(Orig_Path)
                        tmp = New FCS_File(item)
                        mValidRun = tmp.Valid
                        If (mValidRun) Then Exit For
                    Next
                    If (mValidRun) Then
                        If (Orig_Path.ToLower = Data_Path.ToLower) Then

                        End If
                    End If
                End If
            End If
            Return mValidRun
        End Get
    End Property

    ' dgp rev 4/24/09 replace filesystemobject with system.io methods

    ' dgp rev 6/22/07 Hold the name of the first file used for date and checksum
    ' dgp rev 5/28/09 Determine first run file after the BTIM Sort.
    Private mFirstRunFile = Nothing
    Public ReadOnly Property FirstRunFile() As String
        Get
            If mFirstRunFile Is Nothing Then mFirstRunFile = FCS_Files(0).FileName
            Return mFirstRunFile
        End Get
    End Property

    ' dgp rev 8/15/08 Incorporate Exper into Run
    Private mExperFlag As Boolean = False
    Private mBDFACS_Node As XmlNode

    ' dgp rev 8/8/07 Return the experiment doc
    Private xml_exp_file As String
    Private xml_exp_doc As System.Xml.XmlDocument
    Public ReadOnly Property Exp_Doc() As System.Xml.XmlDocument
        Get
            Return xml_exp_doc
        End Get
    End Property

    ' dgp rev 6/20/07 is the XML file a BDFacs
    Public Function Chk_BDFacs(ByVal spec As String) As Boolean

        Chk_BDFacs = False

        If (System.IO.File.Exists(spec)) Then
            If (System.IO.Path.GetExtension(spec).ToLower = ".xml") Then
                Dim test = New System.Xml.XmlDocument
                Dim node As System.Xml.XmlNodeList
                test.Load(spec)
                node = test.SelectNodes("bdfacs")
                If (node.Count > 0) Then
                    Chk_BDFacs = True
                    mBDFACS_Node = node.Item(0)
                    xml_exp_doc = test
                End If
                Chk_BDFacs = (node.Count > 0)
            End If
        End If

    End Function

    ' dgp rev 8/15/08 Real Files
    Private mRealFiles As ArrayList

    ' dgp rev 5/1/09 Swap Bytes
    Private mByteSwapFlag As Boolean = False
    Public Property ByteSwapFlag() As Boolean
        Get
            Return mByteSwapFlag
        End Get
        Set(ByVal value As Boolean)
            mByteSwapFlag = value
        End Set
    End Property

    ' dgp rev 6/20/07 scan the path for XML and FCS files
    Public Function Scan_Path() As Boolean

        mRealFiles = New ArrayList
        Dim item

        For Each item In System.IO.Directory.GetFiles(Me.Orig_Path)
            If (System.IO.Path.GetExtension(item).ToLower = ".xml") Then
                If (Chk_BDFacs(item)) Then xml_exp_file = item
            End If
            If (System.IO.Path.GetExtension(item).ToLower = ".fcs") Then mRealFiles.Add(System.IO.Path.GetFileName(item).ToLower)
        Next

        Scan_Path = mRealFiles.Count > 0 And System.IO.File.Exists(xml_exp_file)

    End Function

    ' dgp rev 6/20/07 List of Files
    Private mDBFiles As ArrayList
    Public Property DBFiles() As ArrayList
        Get
            Return mDBFiles
        End Get
        Set(ByVal value As ArrayList)
            mDBFiles = value
        End Set
    End Property

    '  dgp rev 6/20/07   Private Function File_Names() As Collection
    Public Function Extract_DBFiles() As Boolean

        Dim tmp_node As XmlNode

        DBFiles = New ArrayList
        For Each tmp_node In mBDFACS_Node.SelectNodes("experiment/specimen/tube/data_begin_date")
            Dim name = tmp_node.ParentNode.SelectSingleNode("data_filename")
            If name IsNot Nothing Then DBFiles.Add(name.FirstChild.Value)
        Next

    End Function

    ' dgp rev 8/8/07 Experiment Username
    Private mExperiment As String
    Public ReadOnly Property Experiment() As String
        Get
            Return mExperiment
        End Get
    End Property

    ' Earliest Begin Time
    Private mBTim As Date
    Public Property BTim() As Date
        Get
            Return mBTim
        End Get
        Set(ByVal value As Date)
            mBTim = value
        End Set
    End Property

    ' dgp rev 6/20/07 Run Number
    Private mRun As String
    Public ReadOnly Property Run() As String
        Get
            Return mRun
        End Get
    End Property

    ' dgp rev 8/8/07 Index to first run
    Private mFirstIndex As Integer
    Private mFirst_Specimen As XmlNode

    ' dgp rev 6/20/07 Error Mask
    Private mErrMask As Integer
    Public Property ErrMask() As Integer
        Get
            Return mErrMask
        End Get
        Set(ByVal value As Integer)
            mErrMask = value
        End Set
    End Property

    Const cNoRun As Integer = 1
    '  dgp rev 6/20/07   Private Function File_Names() As Collection
    Public Function Extract_First_Run() As Date

        Dim spc_node, tub_node As XmlNode
        Dim chktime As Date
        Dim node As XmlNodeList
        Dim run As String
        Dim idx As Integer = -1

        Dim RunNum As Integer

        Dim file_list As New Collection

        BTim = Now()
        Dim run_flag As Boolean = False
        For Each spc_node In mBDFACS_Node.SelectNodes("experiment/specimen")
            idx = idx + 1
            run = spc_node.Attributes.Item(0).Value
            Try
                RunNum = CInt(run.Substring(1))
            Catch ex As Exception
                RunNum = 0
            End Try
            If (run.ToLower.StartsWith("r") And RunNum > 0) Then
                run_flag = True
                mRun = run
            End If
            For Each tub_node In spc_node.SelectNodes("tube")
                Try
                    node = tub_node.SelectNodes("date")
                    chktime = node.Item(0).InnerText
                    ' dgp rev 6/22/07 the earliest date with the run
                    If chktime < BTim Then
                        BTim = chktime
                        mFirstIndex = idx
                        mFirst_Specimen = spc_node
                        '                        FirstRunFile = tub_node.Item("data_filename").InnerText
                    End If
                Catch ex As Exception
                End Try
            Next

        Next

        If (Not run_flag) Then mErrMask = mErrMask + cNoRun

    End Function

    Public xml_cksm_doc As System.Xml.XmlDocument

    ' dgp rev 5/19/09 
    Private ReadOnly Property mMDT_UR_Path() As Boolean
        Get
            Dim run_name = New Run_Name(System.IO.Path.GetFileName(mOrig_Path))
            If (run_name.MDT_Flag) Then
                mMDT_UR = run_name.MDTUR
                Return True
            End If
            Return False
        End Get
    End Property

    ' dgp rev 5/19/09 
    Private ReadOnly Property mMDT_UR_Header() As Boolean
        Get
            Return Calc_Run_Name()
        End Get
    End Property
    ' dgp rev 5/19/09 
    Private ReadOnly Property mExtract_MDT_UR() As Boolean
        Get
            ' retrieve from path 
            If (Not mMDT_UR_Path()) Then
                ' dgp rev 5/22/09 
                mRenameFlag = (mDataHomeFlag = True)
                ' else retrieve from header 
                If (Not mMDT_UR_Header()) Then Return False
            End If
            Return True
        End Get
    End Property

    ' dgp rev 5/19/09 
    Private mMDT_UR

    ' dgp rev 5/20/09 
    Public ReadOnly Property MDT_UR_Exists() As Boolean
        Get
            If mMDT_UR Is Nothing Then
                If (mExtract_MDT_UR) Then
                    Return (Not mMDT_UR = "")
                Else
                    Return False
                End If
            Else
                Return (Not mMDT_UR = "")
            End If
        End Get
    End Property
    ' dgp rev 5/22/09 Rename flowroot/data runs not matching standard format
    Private mRenameFlag As Boolean = False

    ' dgp rev 5/22/09 Rename the non-conforming home run
    Private Function RenameRun() As Boolean

        Dim root = System.IO.Directory.GetParent(mOrig_Path).FullName
        Dim target = System.IO.Path.Combine(root, mMDT_UR)
        If Not System.IO.Directory.Exists(target) Then
            System.IO.Directory.Move(mOrig_Path, target)
        End If

    End Function
    ' dgp rev 5/20/09 
    Public ReadOnly Property MDT_UR()
        Get
            If MDT_UR_Exists Then
                If (mRenameFlag) Then
                    '                    RenameRun()
                End If
            End If
            Return mMDT_UR
        End Get
    End Property

    ' dgp rev 6/20/07 Unique Experiment Name
    Public Function Extract_Unique() As String

        Dim dt As String
        Dim tm As String
        Dim sn As String
        Dim node As XmlNodeList

        dt = Format(BTim.Date, "dd-MMM-yyyy")
        tm = Format(BTim, "hh!mm")

        If (Not mBDFACS_Node Is Nothing) Then
            node = mBDFACS_Node.SelectNodes("experiment/specimen/tube/data_instrument_serial_number")
            If (node.Count > 0) Then
                sn = node.Item(0).InnerText()
                Return sn + "_" + Experiment + "_" + Run + "_" + dt + "_" + tm
            End If
        End If

        Return dt + "_" + tm

    End Function

    Private mUniqueFlag As Boolean = False
    Private mUniqueName As String
    Public ReadOnly Property UniqueName() As String
        Get
            If (mUniqueFlag) Then Return mUniqueName
            mUniqueFlag = True
            mUniqueName = Extract_Unique()
            Return mUniqueName
        End Get
    End Property

    ' dgp rev 1/22/09  Original name
    Private mOrigName As String
    Public ReadOnly Property OrigName() As String
        Get
            Return mOrigName
        End Get
    End Property

    Private mFileOrder As ArrayList
    ' dgp rev 3/22/06 create a new FCS Run instance
    Public Sub New(ByVal pathspec As String)

        mInitRun(pathspec)

    End Sub

    ' dgp rev 7/26/07 Mark the checksums for the given run
    ' dgp rev 4/24/09 will be used in upload
    Public Function Mark_Checksums() As Boolean

        Mark_Checksums = True
        Dim cnt As Integer = Me.FCS_cnt
        Dim idx As Integer
        Dim chkstr As String
        Dim chkorg As String
        For idx = 0 To cnt - 1
            chkstr = Me.FCS_Files(idx).ChkSumStr()
            If (Me.FCS_Files(idx).Header.Contains("$Checksum")) Then
                chkorg = Me.FCS_Files(idx).Header("$Checksum")
                If (chkorg.ToLower <> chkstr.ToLower) Then Mark_Checksums = False
            Else
                Me.FCS_Files(idx).Header("$Checksum") = chkstr
            End If
        Next

    End Function


    ' dgp rev 1/17/07 return a unique prefix for the run based upon machine, date, time
    Public Function Unique_Prefix() As String

        If (Run_Flag) Then
            Run_Id = Run_Machine.ToString _
                + "_" + Run_Date.ToString _
                + "_" + Run_Time.ToString _
                + "_" + Run_User.ToString _
                + "_" + Run_Num.ToString
        Else
            Run_Id = Run_Machine.ToString _
             + "_" + Run_Date.ToString _
             + "_" + Run_Time.ToString
        End If

        Return Run_Id

    End Function

    Private Function FindUsers() As Object
        Throw New NotImplementedException
    End Function


End Class
