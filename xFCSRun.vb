' Name:     FCS Run Class
' Author:   Donald G Plugge
' Date:     3/22/06 
' Purpose:  Class to track information concerning a given FCS run
Imports System.io
Imports System.Xml
Imports System.Security.Cryptography
Imports HelperClasses

' DataSet (flowroot\data),  BrowseRun(anywhere)

Public Class FCSRun

    Public Shared objImp As RunAs_Impersonator

    ' use delegates for Repeat
    Delegate Sub delRepeat(ByVal Xname As String)
    Public Event evtRepeat As delRepeat

    Public Good_Backup As Int16
    Public Good_Upload As Int16
    Public Backup_Path As String
    Public Backup_Alias As String
    Public Total_Attempts As Int16

    Private xml_doc As Xml.XmlDocument

    ' dgp rev 4/24/09 Cache the selected dataset
    Public Function SaveToCache() As Boolean

        SaveToCache = False

        Dim root_path As String = System.IO.Path.Combine(FlowStructure.Depot_Root, "FCSRuns")
        Dim exp_path As String = System.IO.Path.Combine(NCIUser, FirstChecksum)
        Dim full_path = System.IO.Path.Combine(root_path, exp_path)
        Try
            If (System.IO.Directory.Exists(full_path)) Then
                For Each item In System.IO.Directory.GetFiles(full_path)
                    System.IO.File.Delete(item)
                Next
            End If
            mStatus = "Cache Deleted"
        Catch ex As Exception
            mStatus = "Cache Delete Failure"
            Exit Function
        End Try
        Try
            If (FCS_cnt > 0) Then
                Dim path
                For idx As Integer = 0 To FCS_cnt - 1
                    path = System.IO.Path.Combine(CachePath, Me.FCS_Files(idx).FileName)
                    FCS_Files(idx).SwapByteFlag = Me.ByteSwapFlag
                    FCS_Files(idx).Save_File(path)
                    SaveToCache = True
                Next
            End If
        Catch ex As Exception
            SaveToCache = False
        End Try
        Return SaveToCache

    End Function

    ' dgp rev 5/1/09 Compare cached data
    Public Function CompareCached() As Boolean

        Dim CacheFile As FCS_File
        Dim OrigFile As FCS_File
        Dim FCScache
        CompareCached = False
        For idx As Integer = 0 To FCS_cnt - 1

            FCScache = System.IO.Path.Combine(DataCache, FCS_List(idx).name)
            CacheFile = New FCS_File(FCScache)
            OrigFile = New FCS_File(FCS_List(idx).fullname)
            If (Not CacheFile.ChkSumStr = OrigFile.ChkSumStr) Then Exit Function

        Next
        CompareCached = True

    End Function

    ' dgp rev 4/20/09 First attempt to work with an instance pointer
    Private mProtocol As FCSTable
    Public Property Protocol() As FCSTable
        Get
            Return mProtocol
        End Get
        Set(ByVal value As FCSTable)
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

    ' dgp rev 4/24/09 Internal merge
    Private Function mMerge(ByVal force As Boolean) As Boolean

        mMerge = False
        If (Protocol Is Nothing) Then Return False

        If (Protocol.myTable.Rows.Count = 0) Then Return False

        If (Not Protocol.myTable.Rows.Count = FCS_cnt And Not force) Then Return False

        Dim FileIdx = 0
        Dim row As DataRow
        For Each row In Protocol.myTable.Rows
            For ColIdx = 0 To row.ItemArray.Length - 1
                If (FCS_Files(FileIdx).Header.ContainsKey(row.Table.Columns.Item(ColIdx))) Then
                    FCS_Files(FileIdx).Header.Add(row.Table.Columns.Item(ColIdx).ColumnName, row.ItemArray(ColIdx))
                Else
                    FCS_Files(FileIdx).Header.Item(row.Table.Columns.Item(ColIdx).ColumnName) = row.ItemArray(ColIdx)
                End If
            Next
            FileIdx = FileIdx + 1
        Next
        mMerge = True

    End Function

    ' dgp rev 4/24/09 Merge only if protocol matches run count
    Public Function Merge() As Boolean

        Return mMerge(False)

    End Function

    ' dgp rev 4/20/09 Attempt to merge run with current protocol
    Public Function Merge(ByVal force As Boolean) As Boolean

        Return mMerge(force)

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
    Private mAssigned As Boolean = False
    Private mUnassignedName
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

    ' dgp rev 7/29/08 valid CurCache of data
    Private mCurCache_Exists As Boolean = False
    Public ReadOnly Property CurCache_Exists() As Boolean
        Get
            Return mCurCache_Exists
        End Get
    End Property

    ' dgp rev 8/6/08 Current Cache
    Private mCur_CurCache As String
    Public ReadOnly Property Cur_CurCache() As String
        Get
            Return mCur_CurCache
        End Get
    End Property

    ' dgp rev 8/9/07 Local depot location for current experiment
    ' assume path indicates data in path
    Private Sub Calc_CurCache()

        If mCurCache_Exists Then Exit Sub
        mCurCache_Exists = False
        If (Valid_Run()) Then
            If (Valid_User()) Then
                Dim root_path As String = System.IO.Path.Combine(FlowStructure.Depot_Root, "FCSRuns")
                Dim exp_path As String = System.IO.Path.Combine(NCIUser, Unique_Prefix)
                mCur_CurCache = System.IO.Path.Combine(root_path, exp_path)
                mCurCache_Exists = System.IO.Directory.Exists(mCur_CurCache)
            End If
        End If

    End Sub

    ' dgp rev 7/12/07 Prefix state 
    Public Enum PrefixType
        [Date_ymd] = 0
        Date_mdy = 1
    End Enum

    ' dgp rev 7/12/07 suffix state 
    Public Enum SuffixType
        [Sample] = 0
        FileOrder = 1
        BTIM = 2
    End Enum

    Public SuffixState As SuffixType = SuffixType.FileOrder
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

        Dim Header As ArrayList

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
        Protocol.ParCnt = Protocol.myTable.Columns.Count
        Protocol.Tubes = Protocol.myTable.Rows.Count
        Return ((Protocol.ParCnt * Protocol.Tubes) > 0)

    End Function

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

        For Each item In System.IO.Directory.GetFiles(Cur_CurCache)
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

        Next

    End Function

    ' prefix using internal $SMNO or TUBE NAME keyword
    Private Function Seq1() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim idx As Integer

        Seq1 = False

        idx = 0
        For Each item In System.IO.Directory.GetFiles(Cur_CurCache)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                idx = idx + 1
                If (objFCS.Header.ContainsKey("$SMNO")) Then
                    val = objFCS.Header("$SMNO")
                Else
                    If (objFCS.Header.ContainsKey("TUBE NAME")) Then
                        val = objFCS.Header("TUBE NAME")
                    Else
                        val = (String.Format("{0:D3}", CInt(idx)))
                    End If
                End If

                RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + val + ".fcs"
                Seq1 = True
            End If
        Next

    End Function

    ' prefix using internal $SMNO or TUBE NAME keyword
    Private Function Seq4() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim count As Integer
        Dim seq As String

        Seq4 = False
        count = 1

        For Each item In System.IO.Directory.GetFiles(Cur_CurCache)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                seq = Format(Val(count), "000")
                count = count + 1
                RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + seq + ".fcs"
                Seq4 = True
            End If
        Next

    End Function

    Private mBTimOrder As ArrayList
    Private mFirstBTim As String
    Private mDate As String
    ' sort by BTim
    Private Function BTimSort() As Boolean

        BTimSort = False

        ' dgp rev 5/26/09 Any Path?
        If (Not mPathExists) Then Exit Function

        ' dgp rev 5/26/09 Any Files?
        If (Not mFilesExist) Then Exit Function

        Dim item
        Dim objFCS As FCS_Classes.FCS_File

        Dim BTim As New Hashtable
        Dim SRT As New ArrayList

        m_fcsfiles = New ArrayList
        Dim unsorted As New Hashtable        

        Dim DateTracker As New ArrayList
        For Each item In System.IO.Directory.GetFiles(mOrig_Path)
            objFCS = New FCS_Classes.FCS_File(item)
            If (objFCS.Valid) Then
                If (objFCS.Header.ContainsKey("$BTIM")) Then
                    unsorted.Add(objFCS.Header("$BTIM"), objFCS)
                    BTim.Add(objFCS.Header("$BTIM"), System.IO.Path.GetFileName(item))
                    SRT.Add(objFCS.Header("$BTIM"))
                End If
                If (objFCS.Header.ContainsKey("$DATE") And Not DateTracker.Contains(objFCS.Header("$DATE"))) _
                Then DateTracker.Add(objFCS.Header("$DATE"))
            End If
        Next

        SRT.Sort()
        mBTimOrder = New ArrayList
        For Each item In SRT
            mBTimOrder.Add(BTim(item))
            m_fcsfiles.Add(unsorted(item))
        Next

        unsorted = Nothing
        BTimSort = (mBTimOrder.Count > 0) And (DateTracker.Count = 1)
        If (BTimSort) Then
            mFirstBTim = SRT(0)
            mDate = DateTracker.Item(0)
        End If

    End Function

    ' prefix using internal $SMNO or TUBE NAME keyword
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

        mMDT_UR = String.Format("{0}_{1}_{2}", SN, Format(Convert.ToDateTime(mDate), "MMM-dd-yy"), Format(Convert.ToDateTime(mFirstBTim), "hh!mm"))

        Return True

    End Function

    ' calculate the new name using selected prefix and sequence
    Public Function Calc_New_Names() As Boolean

        RenameList.Clear()
        Calc_New_Names = False

        Calc_New_Names = Prefix()

        If (Calc_New_Names) Then
            Select Case SuffixState
                Case SuffixType.Sample
                    Calc_New_Names = Seq1()
                Case SuffixType.FileOrder
                    Calc_New_Names = Seq4()
                Case SuffixType.BTIM
                    Calc_New_Names = SeqBTim()
            End Select
        End If

    End Function

    'dgp rev 7/12/07 Rename the Origial Files
    Public Function Rename_Files() As Boolean

        Rename_Files = True

        Dim fil
        Dim new_name As String
        Dim item
        '        File_Count = RenameList.Count
        '       Error_Count = 0
        Dim prev As Integer = RenameList.Count + 1

        Dim Remove_List As New ArrayList

        Dim Proc_List As Dictionary(Of String, String) = RenameList

        ' loop thru the list and remove successful renames
        ' dgp rev 7/11/07 loop inserted to handle repetitive rename for filealreadyexists
        While (Remove_List.Count <> Proc_List.Count)
            Remove_List = New ArrayList
            If (Proc_List.Count = prev) Then Exit While
            prev = Proc_List.Count
            For Each item In Proc_List
                fil = System.IO.Path.GetFileName(System.IO.Path.Combine(Cur_CurCache, item.Key))
                new_name = item.Value
                Try
                    fil.Name = new_name
                    Remove_List.Add(item.Key)
                Catch ex As Exception
                    '                    If (ex.Message.ToLower.Contains("filealreadyexists")) Then flgSpecCase = True
                    Rename_Files = False
                    '                    Error_Count = Error_Count + 1
                End Try
            Next
            Dim rem_item
            For Each rem_item In Remove_List
                Proc_List.Remove(rem_item)
            Next
        End While
        If (Proc_List.Count = 0) Then Rename_Files = True

    End Function

    ' dgp rev 7/30/08 NCI User name
    Private mNCIUser As String
    Public Property NCIUser() As String
        Get
            Return mNCIUser
        End Get
        Set(ByVal value As String)
            mNCIUser = value
            ' dpg rev 8/6/08 once user is set, cache can be calculated
            Calc_CurCache()
            ' dgp rev 7/30/08 Experiment User is selected
            If (CurStage = Stage.Nothing) Then Prep()
        End Set
    End Property

    Public Run_List_Path As String = "\\Nt-eib-10-6b16\FTP_root\runs"

    ' dgp rev 7/29/08 Valid User consists of Valid Experiment and User from List
    Public Function Valid_User() As Boolean

        If (Not Valid_Run()) Then Return False
        If (UserList Is Nothing) Then Return False
        Return UserList.Contains(NCIUser)

    End Function

    ' dgp rev 7/29/08 Valid User only if contained within UserList
    Private mUserList As ArrayList
    Public Property UserList() As ArrayList
        Get
            Return mUserList
        End Get
        Set(ByVal value As ArrayList)
            mUserList = value
        End Set
    End Property

    ' dgp rev 10/16/08
    Public ReadOnly Property Successful_Upload() As Boolean
        Get
            Return (Upload_Info.TransferCount = Upload_Info.SuccessfulCount)
        End Get
    End Property

    ' dpg rev 10/16/08 check the upload status and if successful, clear cache
    Public Function CheckNClear() As Boolean

        CheckNClear = False
        If (CurStage = FCSUpload.Stage.Uploaded) Then Return True
        If (Successful_Upload) Then
            If (System.IO.Directory.Exists(Cur_CurCache)) Then
                Try
                    System.IO.Directory.Delete(Cur_CurCache, True)
                    CheckNClear = True
                    CurStage = FCSUpload.Stage.Uploaded
                    mStatus = "Successful Upload"
                Catch ex As Exception
                    mStatus = "Cache removal error"
                End Try
            End If
        End If

    End Function


    ' dgp rev 7/11/07 upload data to server
    ' may need to use authentication
    ' dgp rev 7/27/07 Upload an FCS Run to Server
    ' dgp rev 3/4/09 Upload looks for the RESERVED run to replace
    ' dgp rev 3/5/09 Look for empty reserved folder and replace with data folder
    ' dgp rev 3/5/09 separate the building of the source and target from the transfer
    ' dgp rev rev 3/10/09 move upload routine to FCSUpload class
    Public Enum Stage
        Invalid = -1
        Mismatch = -2
        [Nothing] = 0
        Sync = 1
        Ready = 2
        Uploaded = 3
        Stored = 4
    End Enum

    ' dgp rev 11/3/08 Color of Current Stage
    Public ReadOnly Property StageColor() As System.Drawing.Color
        Get
            Select Case mCurStage
                Case Stage.Invalid
                    Return Drawing.Color.Red
                Case Stage.Mismatch
                    Return Drawing.Color.Red
                Case Stage.Nothing
                    Return Drawing.Color.White
                Case Stage.Ready
                    Return Drawing.Color.GreenYellow
                Case Stage.Uploaded
                    Return Drawing.Color.Green
                Case Stage.Sync
                    Return Drawing.Color.Aqua
                Case Stage.Stored
                    Return Drawing.Color.LightGreen
            End Select

        End Get
    End Property

    ' dgp rev 7/29/08 What is the current stage of the transfer
    ' none, ready or done -- errors include invalid and mismatch
    Private mCurStage As Stage = Stage.Nothing
    Public Property CurStage() As Stage
        Get
            Return mCurStage
        End Get
        Set(ByVal value As Stage)
            mCurStage = value
        End Set
    End Property

    ' dgp rev 6/22/07 checksum of first file to uniquely ident run or experiment
    Private mFirstChecksum As String = "X"

    Public ReadOnly Property CachePath() As String
        Get
            Return System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.FlowRoot, "LocalCache"), mFirstChecksum)
        End Get
    End Property

    Public ReadOnly Property FirstChecksum() As String
        Get
            Return mFirstChecksum
        End Get
    End Property
    ' 8/6/08 Checksum Exists
    Public ReadOnly Property ChkSum_Exists() As Boolean
        Get
            If (Valid_Run()) Then Return System.IO.File.Exists(System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums"), mFirstChecksum + ".xml"))
            Return False
        End Get
    End Property

    ' dgp rev 8/6/08 Checksum Path, may or may not exist
    Public ReadOnly Property ChkSum_Path() As String
        Get
            If (Valid_Run()) Then Return System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums"), mFirstChecksum + ".xml")
            Return ""
        End Get
    End Property

    ' dgp rev 8/6/08 
    Private Function CheckServer() As Boolean

        Dim Server_Path As String
        CheckServer = False

        If (Not Upload_Info Is Nothing) Then

            Server_Path = String.Format("\\{0}{1}", Upload_Info.DataServer, Upload_Info.Upload_Root)
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

    ' dgp rev 8/1/08
    Public Function Store_Away() As Boolean

        Select Case CurType
            Case FCSType.Experiment
                If (IsUploaded) Then
                    Store_Away = Upload_Info.Store_Exper()
                    mStoredFlag = Store_Away
                    If (Remove_Original()) Then

                    End If
                End If
            Case FCSType.Run
                If (IsUploaded) Then
                    Store_Away = Upload_Info.Store_Files()
                    mStoredFlag = Store_Away
                End If
        End Select

    End Function

    ' dgp rev 8/6/08 calculate what stage the data distribution is in.
    Public Sub Calc_Stage()

        ' default to invalid and prove otherwise
        CurStage = FCSRun.Stage.Invalid
        ' valid FCS path -- at least one FCS file
        If (Valid_Run()) Then
            CurStage = FCSRun.Stage.Nothing
            ' nothing can be done until a user is linked to the data
            If (Valid_User()) Then
                ' sync the data 
                CurStage = FCSRun.Stage.Sync
                If (ChkSum_Exists) Then
                    CurStage = FCSRun.Stage.Ready
                    If (Not mCurCache_Exists) Then
                        If (CheckServer()) Then
                            CurStage = FCSRun.Stage.Uploaded
                        Else
                            CurStage = FCSRun.Stage.Sync
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private mStatus As String

    ' dgp rev 8/6/08 Original Run Path
    Private mOrig_Path As String
    Private mValidRun

    Private Sub hold()

        Dim di As DirectoryInfo
        Dim rn As Run_Name
        Dim pathspec As String

        mOrigName = System.IO.Path.GetDirectoryName(pathspec)
        If (Me.Scan_Path()) Then
            Validate_Exp()
        End If

        rn = New Run_Name(di.Name.ToString)

        Dim item As FileInfo
        Dim objFCS As FCS_File

        Dim ChkTime As Date
        Dim BTimHash As New Hashtable
        Dim BTimList As New ArrayList

        mFileOrder = New ArrayList
        For Each item In FCS_List
            objFCS = New FCS_File(item.FullName.ToString())
            If (objFCS.Valid) Then
                m_fcsfiles.Add(objFCS)
                If (objFCS.Header.ContainsKey("$BTIM")) Then
                    ChkTime = objFCS.Header.Item("$BTIM")
                    If (Not BTimHash.ContainsKey(ChkTime.TimeOfDay.Ticks)) Then
                        BTimHash.Add(ChkTime.TimeOfDay.Ticks, item.Name.ToUpper)
                        BTimList.Add(ChkTime.TimeOfDay.Ticks)
                    End If
                End If
            End If
            'Console.WriteLine("Par count " + objFCS.Header("$PAR"))
        Next

        ' dgp rev 4/24/09 sort based upon collection time ($BTIM)
        BTimList.Sort()
        Dim entry
        For Each entry In BTimList
            mFileOrder.Add(BTimHash(entry))
        Next

        FirstRunFile = mFileOrder.Item(0)
        Calc_ChkSm_All()
        ' xyzzy User and Run Number should agree with next available run
        IsAria()

        mDynRun = New Dynamic(DataInfoRoot, FirstChecksum)


    End Sub

    Private mValidPath As Boolean = False
    Private mDirInfo As DirectoryInfo
    Private mDataHomeFlag As Boolean = False
    Private mDataCachedFlag As Boolean = False

    Private mPathExists As Boolean = False
    Private mFilesExist As Boolean = False

    ' dgp rev 5/19/09 
    Private Sub mInitRun(ByVal pathspec As String)

        mOrig_Path = pathspec
        ' dgp rev 5/19/09 at a minimum, path must exist and contain files
        If (Not System.IO.Directory.Exists(pathspec)) Then Return
        mPathExists = True
        If (System.IO.Directory.GetFiles(pathspec).Length = 0) Then Return
        mFilesExist = True
        mDirInfo = New DirectoryInfo(pathspec)
        If (mDirInfo.Parent.Name.ToLower = "data") Then If (mDirInfo.Parent.Parent.Name.ToLower = "flowroot") Then mDataHomeFlag = True
        If (mDirInfo.Parent.Parent.Name.ToLower = "localcache") Then If (mDirInfo.Parent.Parent.Parent.Name.ToLower = "flowroot") Then mDataCachedFlag = True

        m_FCS_cnt = Nothing
        mValidRun = Nothing
        mValidPath = True

    End Sub

    Public ReadOnly Property Orig_Path() As String
        Get
            Return mOrig_Path
        End Get
    End Property

    Private mFile_ChkSum_List As ArrayList

    ' dgp rev 6/26/07 Calculate single checksum of FCS binary data block
    Public Function Calc_FCS_Chksum(ByVal file As String) As String

        If (Not System.IO.File.Exists(file)) Then Return ""

        Dim objFCS As FCS_Classes.FCS_File

        mFile_ChkSum_List = New ArrayList
        objFCS = New FCS_Classes.FCS_File(file)
        Calc_FCS_Chksum = objFCS.ChkSumStr

    End Function


    ' dgp rev 4/5/07 Save the Checksum in XML format
    ' dgp rev 6/25/07 Save Checksum for each file
    Public Function Save_Checksum() As Boolean

        Save_Checksum = False

        Dim XML_DataSet As New DataSet("Validator")
        Dim XML_Table As New DataTable("List")

        Dim idx As Int16
        Dim nc As DataColumn
        Dim item

        For Each item In FCS_List
            nc = New DataColumn
            nc.ColumnName = item.Name.ToLower.Replace(".fcs", "")
            '            nc.Caption = item.FullName.ToString()
            XML_Table.Columns.Add(nc)
        Next

        Dim myRow As DataRow

        ' fill in the rows
        myRow = XML_Table.NewRow()
        ' dgp rev 6/25/07 
        For idx = 0 To mFile_ChkSum_List.Count - 1
            ' Then add the new row to the collection.
            myRow(idx) = mFile_ChkSum_List(idx).ToString
        Next
        XML_Table.Rows.Add(myRow)
        XML_DataSet.Tables.Add(XML_Table)
        If (Utility.Create_Tree(System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums"))) Then
            Try
                XML_DataSet.WriteXml(ChkSum_Path, XmlWriteMode.WriteSchema)
                Save_Checksum = True
            Catch ex As Exception

            End Try
        End If

    End Function

    ' dgp rev 4/5/07 Save the Checksum in XML format
    ' dgp rev 6/25/07 Save Checksum for each file
    Public Function Save_DataInfo() As Boolean

        Save_DataInfo = False

        Dim XML_DataSet As New DataSet("Validator")
        Dim XML_Table As New DataTable("List")

        Dim idx As Int16
        Dim nc As DataColumn
        Dim item

        For Each item In FCS_List
            nc = New DataColumn
            nc.ColumnName = item.Name.ToLower.Replace(".fcs", "")
            '            nc.Caption = item.FullName.ToString()
            XML_Table.Columns.Add(nc)
        Next

        Dim myRow As DataRow

        ' fill in the rows
        myRow = XML_Table.NewRow()
        ' dgp rev 6/25/07 
        For idx = 0 To mFile_ChkSum_List.Count - 1
            ' Then add the new row to the collection.
            myRow(idx) = mFile_ChkSum_List(idx).ToString
        Next
        XML_Table.Rows.Add(myRow)
        XML_DataSet.Tables.Add(XML_Table)
        If (Utility.Create_Tree(DataInfoPath)) Then
            Try
                XML_DataSet.WriteXml(System.IO.Path.Combine(DataInfoPath, mFirstChecksum), XmlWriteMode.WriteSchema)
                Save_DataInfo = True
            Catch ex As Exception

            End Try
        End If

    End Function

    Private mDataInfoPath

    ' dgp rev 4/29/09 
    Public Function DataInfoPath() As String

        If (mDataInfoPath Is Nothing) Then
            If (Check_DataInfo()) Then Return mDataInfoPath
        End If
        Return ""

    End Function

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

    ' dgp rev 4/29/09 Data Info Path Flag
    Private mDataInfoPathFlag
    ' dgp rev 4/29/09 Check for Data Info
    Public Function Check_DataInfo() As Boolean

        If (mDataInfoPathFlag Is Nothing) Then
            mDataInfoPath = System.IO.Path.Combine(DataInfoRoot, FirstChecksum)
            mDataInfoPathFlag = System.IO.File.Exists(mDataInfoPath)
        End If
        Return mDataInfoPathFlag

    End Function

    ' dgp rev 7/11/07 create the CurCache
    Private Function Create_CurCache() As Boolean

        ' dgp rev 7/29/08 in order to create the CurCache we need a valid experiment
        ' and a valid root, but we've already determined this, right?  xyzzy
        If (Valid_Run() And Valid_User()) Then
            If (Not mCurCache_Exists) Then
                mCurCache_Exists = Utility.Create_Tree(mCur_CurCache)
            End If
        End If
        Return mCurCache_Exists

    End Function

    ' dgp rev 6/26/07 Change the Experiment Location
    ' move the selected experiment into a local cache in the flowroot depot
    ' then add a checksum for validation
    Private Function Move_To_Cache() As Boolean

        Move_To_Cache = False
        mStatus = "Directory Failure"

        If (Valid_Run() And Valid_User()) Then
            If (Not Create_CurCache()) Then Exit Function
            Try
                System.IO.File.Copy(Orig_Path, mCur_CurCache)
            Catch ex As Exception
                mStatus = "Move Failure"
                Exit Function
            End Try
            If (Save_Checksum()) Then
                CurStage = FCSUpload.Stage.Ready
                mStatus = "Ready for Upload"
            Else
                mStatus = "Checksum Failure"
            End If
        End If

        mStatus = "Successful Move"
        Move_To_Cache = True

    End Function

    Private mDataInfoRoot

    ' dgp rev 4/29/09 Data Inforamation - checksum list, user, run
    Private Function DataInfoRoot()

        If (mDataInfoRoot Is Nothing) Then
            mDataInfoRoot = System.IO.Path.Combine(FlowStructure.RemoteRoot, "Settings")
            mDataInfoRoot = System.IO.Path.Combine(mDataInfoRoot, "Data")
        End If
        Return mDataInfoRoot

    End Function

    Private mXML_ChkSum_List As ArrayList
    ' dgp rev 6/26/07 
    Private Function Load_XML_ChkSum() As Boolean

        Load_XML_ChkSum = False

        If (Not ChkSum_Exists) Then Exit Function

        '        Log_Info("Loading " + ChkSum_Path)
        xml_doc = New Xml.XmlDocument
        xml_doc.Load(ChkSum_Path)

        Dim nodes As Xml.XmlNodeList
        Dim node As Xml.XmlNode

        Try
            '           Log_Info("Scan for Nodes")
            nodes = xml_doc.SelectNodes("Checksum_x0020_Validator/Checksum_x0020_List")
            If (nodes.Count = 0) Then nodes = xml_doc.SelectNodes("Validator/List")
        Catch ex As Exception
            '          Log_Info("Node Scan Error")
            Exit Function
        End Try

        If (nodes.Count = 0) Then
            '         Log_Info("Node Not Found")
        Else
            '            Log_Info("Node Found")
            ' check sum list read from XML file
            ' must compare to checksums from actual files
            mXML_ChkSum_List = New ArrayList
            For Each node In nodes.Item(0).ChildNodes
                mXML_ChkSum_List.Add(node.InnerText)
            Next
            If (mXML_ChkSum_List.Count > 1) Then Load_XML_ChkSum = True
        End If

    End Function

    ' dgp rev 6/26/07 Compare the Actual File List with the saved XML list
    Public Function Compare_Lists() As Boolean

        Compare_Lists = False
        mStatus = "Actual Checksum Error"
        If (mFile_ChkSum_List Is Nothing) Then Exit Function
        mStatus = "XML Checksum Error"
        If (mXML_ChkSum_List Is Nothing) Then Exit Function
        mStatus = "File Count Mismatch"
        If (mXML_ChkSum_List.Count <> mFile_ChkSum_List.Count) Then Exit Function

        mFile_ChkSum_List.Sort()
        mXML_ChkSum_List.Sort()

        Dim idx
        mStatus = "Checksum Mismatch"
        For idx = 0 To mXML_ChkSum_List.Count - 1
            If (Not mFile_ChkSum_List.Contains(mXML_ChkSum_List.Item(idx))) Then Exit Function
        Next

        mStatus = "Valid Checksum"
        Return True

    End Function


    ' dgp rev 6/26/07 Verify that the checksum matches the file checksum
    Private Function Verify_Checksum() As Boolean

        Verify_Checksum = False

        If (Load_XML_ChkSum()) Then
            Return Compare_Lists()
        Else
            mStatus = "No XML checksum file"
        End If

        Return False

    End Function

    ' dgp rev 6/25/07 Prep experiment
    ' errors - mismatch, caching, or validation
    Public Function Prep() As Boolean

        Calc_Stage()
        If (CurStage = Stage.Nothing) Then Exit Function
        If (CurStage = Stage.Sync) Then
            ' dgp rev 7/29/08 experiment already in cache, simple validate
            ' dgp rev 7/29/08 Compare Current Experiment with Existing Checksum
            If (Move_To_Cache()) Then CurStage = FCSUpload.Stage.Ready
        ElseIf (CurStage = Stage.Ready) Then
            ' dgp rev 7/29/08 experiment needs to be cached and a checksum calculated
            If (Not Verify_Checksum()) Then CurStage = FCSUpload.Stage.Mismatch
        End If
        Calc_Stage()

        Return (CurStage = Stage.Nothing Or CurStage = Stage.Ready Or CurStage = Stage.Uploaded)

    End Function

    ' dgp rev 10/12/07 Backup FCS Data
    Public Sub Backup()

        Good_Backup = 0
        Backup_Path = String.Format("\\ncifs-p007.nci.nih.gov\Group\EIB\Branch99\FCSAria\")
        Backup_Path = System.IO.Path.Combine(Backup_Path, Upload_Info.AssignedUser)

        Backup_Path = System.IO.Path.Combine(Backup_Path, Upload_Info.Full_Name)
        Backup_Alias = Backup_Path.Replace("\\ncifs-p007.nci.nih.gov\Group\EIB\Branch99", "I:")

        If (System.IO.Directory.Exists(Backup_Path)) Then
            If (MsgBox("Backup already exists.  Overwrite?", MsgBoxStyle.YesNo) = MsgBoxResult.No) Then Exit Sub
        End If
        Dim SrcFil
        Dim DestFil As String
        Total_Attempts = System.IO.Directory.GetFiles(Cur_CurCache).Length
        Utility.Create_Tree(Backup_Path)

        RaiseEvent evtRepeat("Backup...")
        For Each SrcFil In System.IO.Directory.GetFiles(Cur_CurCache)
            RaiseEvent evtRepeat(System.IO.Path.GetFileName(SrcFil))
            DestFil = System.IO.Path.Combine(Backup_Path, System.IO.Path.GetFileName(SrcFil))
            System.IO.File.Copy(SrcFil, DestFil, True)
            If System.IO.File.Exists(DestFil) Then Good_Backup = Good_Backup + 1
        Next

    End Sub


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

    Private m_chksum_list As New Collection
    Public Property ChkSum_List() As Collection
        Get
            Return m_chksum_list
        End Get
        Set(ByVal value As Collection)
            m_chksum_list = value
        End Set
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

    ' array of FCS files objects
    Private m_fcsfiles As ArrayList = Nothing
    Public Property FCS_Files(ByVal idx As Int16) As FCS_File
        Get
            If m_fcsfiles Is Nothing Then If Not BTimSort() Then Return Nothing
            If m_fcsfiles Is Nothing Then Return Nothing
            If idx > m_fcsfiles.Count - 1 Then Return Nothing
            Return m_fcsfiles(idx)
        End Get
        Set(ByVal Value As FCS_File)
            m_fcsfiles.Add(Value)
        End Set
    End Property

    ' dgp rev 5/9/06 Calculate Checksum
    Public Sub Calc_ChkSum_List()

        Dim objFCS As FCS_File
        Dim item As FileInfo
        Dim chksum As Byte()
        Dim bite As Byte
        Dim HEXStr As String

        ChkSum_List.Clear()

        For Each item In FCS_List
            If (LCase(item.Extension) = ".fcs") Then
                HEXStr = ""
                objFCS = New FCS_File(item.FullName.ToString)
                chksum = objFCS.Calc_ChkSum
                For Each bite In chksum
                    HEXStr = HEXStr + (String.Format("{0:X2}", bite))
                Next
                ChkSum_List.Add(HEXStr.ToString(), item.Name)
            End If
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
                If (mDirInfo IsNot Nothing) Then
                    ' are any of the files proper FCS?
                    Dim item
                    Dim tmp As FCS_File
                    For Each item In System.IO.Directory.GetFiles(mOrig_Path)
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

    Private mDataCache As String
    Private mCacheFlag As Boolean = False

    ' dgp rev 4/24/09 Data has been cached
    Public ReadOnly Property CacheFlag() As Boolean
        Get
            Return mCacheFlag
        End Get
    End Property

    Public ReadOnly Property DataCache() As String
        Get
            If mDataCache Is Nothing Then
                Dim source = New System.IO.DirectoryInfo(Orig_Path)
                Dim rootpath = System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.FlowRoot, "LocalCache"), mFirstChecksum)
                Dim dest = New System.IO.DirectoryInfo(rootpath)
                If (dest.Exists) Then dest.Delete(True)
                Dim newpath = System.IO.Path.Combine(rootpath, source.Name)
                dest = New System.IO.DirectoryInfo(newpath)
                dest.Create()
                mDataCache = dest.FullName
            End If
            Return mDataCache
        End Get
    End Property

    ' dgp rev 4/24/09 Cache the selected dataset
    Public Function CacheCurrent() As Boolean

        Try
            If (FCS_cnt > 0) Then
                Dim path
                For idx As Integer = 0 To FCS_cnt - 1
                    path = System.IO.Path.Combine(DataCache, Me.FCS_Files(idx).FileName)
                    FCS_Files(idx).SwapByteFlag = Me.ByteSwapFlag
                    Me.FCS_Files(idx).Save_File(path)
                Next
            End If
            mCacheFlag = True
        Catch ex As Exception
            mCacheFlag = False
        End Try
        Return mCacheFlag

    End Function

    ' dgp rev 4/24/09 Cache the selected dataset
    Public Function CopytoCache() As Boolean

        Try
            Dim source = New System.IO.DirectoryInfo(Orig_Path)
            Dim newpath = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.Combine(FlowStructure.FlowRoot, "LocalCache"), mFirstChecksum), source.Name)
            If (source.GetFiles.Length > 0) Then
                Dim dest = New System.IO.DirectoryInfo(newpath)
                If (Not dest.Exists) Then dest.Create()
                mDataCache = dest.FullName
                For Each SourceFile As FileInfo In source.GetFiles
                    SourceFile.CopyTo(System.IO.Path.Combine(dest.FullName, SourceFile.Name), True)
                Next
            End If
            mCacheFlag = True
        Catch ex As Exception
            mCacheFlag = False
        End Try
        Return mCacheFlag

    End Function

    ' dgp rev 4/24/09 replace filesystemobject with system.io methods

    ' dgp rev 6/22/07 Hold the name of the first file used for date and checksum
    Private mFirstRunFile As String
    Public Property FirstRunFile() As String
        Get
            Return mFirstRunFile
        End Get
        Set(ByVal value As String)
            mFirstRunFile = value
            ' calculate the checksum when first file is set
            Dim fcs_file As String = System.IO.Path.Combine(Orig_Path, mFirstRunFile)
            mFirstChecksum = Calc_FCS_Chksum(fcs_file)
        End Set
    End Property

    ' dgp rev 6/26/07 Fill File Checksum List with Calculated Checksums
    Public Sub Calc_ChkSm_All()

        Dim item
        Dim objFCS As FCS_Classes.FCS_File

        mFile_ChkSum_List = New ArrayList
        For Each item In FCS_List
            objFCS = New FCS_Classes.FCS_File(item.fullname)
            mFile_ChkSum_List.Add(objFCS.ChkSumStr)
        Next

    End Sub

    ' dgp rev 8/15/08 Incorporate Exper into Run
    Private mExperFlag As Boolean = False
    Private mBDFACS_Node As XmlNode

    ' dgp rev 8/8/07 Return the experiment doc
    Private xml_exp_file As String
    Private xml_exp_doc As Xml.XmlDocument
    Public ReadOnly Property Exp_Doc() As Xml.XmlDocument
        Get
            Return xml_exp_doc
        End Get
    End Property

    ' dgp rev 6/20/07 is the XML file a BDFacs
    Public Function Chk_BDFacs(ByVal spec As String) As Boolean

        Chk_BDFacs = False

        If (System.IO.File.Exists(spec)) Then
            If (System.IO.Path.GetExtension(spec).ToLower = "xml") Then
                Dim test = New Xml.XmlDocument
                Dim node As Xml.XmlNodeList
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
    Private mRealFiles As New ArrayList
    Public Property RealFiles() As ArrayList
        Get
            Return mRealFiles
        End Get
        Set(ByVal value As ArrayList)
            mRealFiles = value
        End Set
    End Property

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

        For Each item In System.IO.Directory.GetFiles(Me.Orig_Path)
            If (System.IO.Path.GetExtension(item).ToLower = "xml") Then
                If (Chk_BDFacs(item)) Then xml_exp_file = item
            End If
            If (System.IO.Path.GetExtension(item).ToLower = ".fcs") Then RealFiles.Add(item)
        Next

        Scan_Path = System.IO.File.Exists(xml_exp_file)

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
        For Each tmp_node In mBDFACS_Node.SelectNodes("experiment/specimen/tube/data_filename")
            DBFiles.Add(tmp_node.FirstChild.Value)
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
                For Each tub_node In spc_node.SelectNodes("tube")
                    node = tub_node.SelectNodes("date")
                    chktime = node.Item(0).InnerText
                    ' dgp rev 6/22/07 the earliest date with the run
                    If chktime < BTim Then
                        BTim = chktime
                        mFirstIndex = idx
                        mFirst_Specimen = spc_node
                        FirstRunFile = tub_node.Item("data_filename").InnerText
                    End If
                Next
            End If
        Next

        If (Not run_flag) Then mErrMask = mErrMask + cNoRun

    End Function

    Public Enum FCSType
        Invalid = -1
        [Nothing] = 0
        Experiment = 1
        Run = 2
    End Enum

    Public CurType As FCSType = FCSType.Nothing

    ' dgp rev 6/20/07 Is this a valid experiment
    Private mValid_Exper As Boolean = False
    Public ReadOnly Property Valid_Exper() As Boolean
        Get
            Return mValid_Exper
        End Get
    End Property


    ' dgp rev 6/19/07 Validate Experiment XML and FCS files
    ' dgp rev 12/14/07 Add a status member to record validation results
    Public Function Validate_Exp() As Boolean

        Dim tube_cnt As Integer = 0
        Dim tmp_list As New ArrayList
        Dim missing As New ArrayList

        Dim btim As Date = Now()

        Validate_Exp = False

        ' scan all the FCS files in the path
        Dim fil
        For Each fil In System.IO.Directory.GetFiles(Orig_Path)
            If (System.IO.Path.GetExtension(fil).ToLower = "fcs") Then tmp_list.Add(System.IO.Path.GetFileName(fil))
        Next
        ' exit if no external FCS files
        If tmp_list.Count = 0 Then
            mStatus = "No FCS files"
            Exit Function
        End If

        Dim item
        Extract_DBFiles()
        For Each item In DBFiles
            If (tmp_list.Contains(item)) Then
                ' remove valid files from list
                tmp_list.RemoveAt(tmp_list.IndexOf(item))
            Else
                ' add a missing entry if not found
                missing.Add(item)
            End If
        Next
        ' exit if no valid XML experiment file
        If (mBDFACS_Node Is Nothing) Then
            mStatus = "No XML experiment file"
            Exit Function
        End If

        mExperiment = mBDFACS_Node("experiment").Attributes.GetNamedItem("name").Value
        Extract_First_Run()

        ' make sure no tubes are missing from XML and 
        ' that no actual FCS files are missing from path
        mValid_Exper = (missing.Count = 0 And tmp_list.Count = 0)

        If (mValid_Exper) Then
            mStatus = "XML Validated"
        Else
            If (missing.Count > 0) Then
                mStatus = "Missing XML entries"
            Else
                mStatus = "Missing FCS files"
            End If
        End If

        Calc_ChkSm_All()

        Return mValid_Exper

    End Function
    ' dgp rev 7/31/08 Experiment Ready for Upload
    Public ReadOnly Property IsReady() As Boolean
        Get
            'Calc_Stage()
            Return (CurStage = Stage.Ready)
        End Get
    End Property

    ' dgp rev 7/31/08 Experiment Ready for Upload
    Public ReadOnly Property IsUploaded() As Boolean
        Get
            'Calc_Stage()
            Return (CurStage = Stage.Uploaded)
        End Get
    End Property

    Public xml_cksm_doc As Xml.XmlDocument
    ' dgp rev 7/11/07 remove the cache
    Private mDataRemoved As Boolean = False
    Public ReadOnly Property DataRemoved() As Boolean
        Get
            Return mDataRemoved
        End Get
    End Property

    ' dgp rev 3/11/09 Remove CurRun, perhaps a valid backup must be confirmed
    Public Function Remove_Original() As Boolean

        ' dgp rev 7/29/08 in order to create the CurCache we need a valid experiment
        ' and a valid root
        Remove_Original = False
        If (Valid_Exper Or Valid_Run()) Then
            Remove_Original = True
            mDataRemoved = True
            Try
                System.IO.Directory.Delete(Orig_Path)
            Catch ex As Exception
                Remove_Original = False
            End Try
        End If

    End Function

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
        System.IO.Directory.Move(mOrig_Path, target)

    End Function
    ' dgp rev 5/20/09 
    Public ReadOnly Property MDT_UR()
        Get
            If MDT_UR_Exists Then
                If (mRenameFlag) Then
                    RenameRun()
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

End Class
