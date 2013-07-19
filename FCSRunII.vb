' Name:     FCS Run Class
' Author:   Donald G Plugge
' Date:     3/22/06 
' Purpose:  Class to track information concerning a given FCS run
Imports System.io
Imports System.Xml
Imports System.Security.Cryptography
Imports HelperClasses

' DataSet (flowroot\data),  BrowseRun(anywhere)

Public Class FCSRunII

    Public Shared objImp As RunAs_Impersonator

    ' use delegates for Repeat
    Delegate Sub delRepeat(ByVal Xname As String)
    Public Event evtRepeat As delRepeat

    ' use delegates for Repeat
    Delegate Sub RenameEventHandler(ByVal OldName As String, ByVal NewName As String)
    Public Event RenameEvent As RenameEventHandler

    Public Good_Backup As Int16
    Public Good_Upload As Int16
    Public Backup_Path As String
    Public Backup_Alias As String
    Public Total_Attempts As Int16

    Private xml_doc As Xml.XmlDocument

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

    ' dgp rev 4/24/09 Cache the selected dataset
    ' dgp rev 9/30/09 Save the checksum path
    Public Function CacheRun() As Boolean

        CacheRun = False
        If (Not CachePathExists()) Then
            Try
                If (FCS_cnt > 0) Then
                    Dim path
                    For idx As Integer = 0 To FCS_cnt - 1
                        path = System.IO.Path.Combine(CachePath, Me.FCS_Files(idx).FileName)
                        FCS_Files(idx).SwapByteFlag = Me.ByteSwapFlag
                        FCS_Files(idx).Save_File(path)
                        CacheRun = True
                    Next
                End If
            Catch ex As Exception
                CacheRun = False
            End Try
        End If
        Return CacheRun

    End Function

    ' dgp rev 4/24/09 Internal merge
    Private Function mMerge(ByVal force As Boolean) As Boolean

        mMerge = False
        If (Protocol Is Nothing) Then Return False

        If (Protocol.myTable.Rows.Count = 0) Then Return False

        If (Not Protocol.myTable.Rows.Count = FCS_cnt And Not force) Then Return False

        Dim FileIdx = 0
        Dim row As DataRow
        Dim colidx
        For Each row In Protocol.myTable.Rows
            For colidx = 0 To row.ItemArray.Length - 1
                If (FCS_Files(FileIdx).Header.ContainsKey(row.Table.Columns.Item(colidx))) Then
                    FCS_Files(FileIdx).Header.Add(row.Table.Columns.Item(colidx).ColumnName, row.ItemArray(colidx))
                Else
                    FCS_Files(FileIdx).Header.Item(row.Table.Columns.Item(colidx).ColumnName) = row.ItemArray(colidx)
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

        If (CachePathExists()) Then
            For Each item In System.IO.Directory.GetFiles(CachePath)
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
        End If

    End Function

    ' prefix using internal $SMNO or TUBE NAME keyword
    Private Function Seq1() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim val As String
        Dim idx As Integer

        Seq1 = False

        idx = 0

        If (CachePathExists()) Then
            For Each item In System.IO.Directory.GetFiles(CachePath)
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
        End If

    End Function

    ' prefix using internal $SMNO or TUBE NAME keyword
    Private Function Seq4() As Boolean

        Dim item
        Dim objFCS As FCS_Classes.FCS_File
        Dim count As Integer
        Dim seq As String

        Seq4 = False
        count = 1

        If (CachePathExists()) Then
            For Each item In System.IO.Directory.GetFiles(CachePath)
                objFCS = New FCS_Classes.FCS_File(item)
                If (objFCS.Valid) Then
                    seq = Format(Val(count), "000")
                    count = count + 1
                    RenameList.Item(System.IO.Path.GetFileName(item)) = RenameList.Item(System.IO.Path.GetFileName(item)) + seq + ".fcs"
                    Seq4 = True
                End If
            Next
        End If

    End Function

    Private mBTimOrder As ArrayList
    Private mFirstBTim = Nothing

    Public ReadOnly Property FirstBTim() As String
        Get
            If (mFirstBTim Is Nothing) Then BTimSort()
            Return mFirstBTim
        End Get
    End Property

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

        m_FCS_cnt = m_fcsfiles.Count
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

        mMDT_UR = String.Format("{0}_{1}_{2}", SN, Format(Convert.ToDateTime(mDate), "MMM-dd-yy"), Format(Convert.ToDateTime(FirstBTim), "hh!mm"))

        Return True

    End Function

    Private mRenameList = Nothing


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
        If (CachePathExists()) Then
            mRenameCount = 0
            mRenameErrors = 0
            While (Remove_List.Count <> Proc_List.Count)
                Remove_List = New ArrayList
                If (Proc_List.Count = prev) Then Exit While
                prev = Proc_List.Count
                For Each item In Proc_List
                    fil = System.IO.Path.Combine(CachePath, item.Key)
                    new_name = System.IO.Path.Combine(CachePath, item.Value)
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
            If (Proc_List.Count = 0) Then Rename_Files = True
        End If

    End Function

    ' dgp rev 10/1/09 Run Cache
    Private mValidCache = Nothing
    Private mCachePath = Nothing

    Public ReadOnly Property ValidCache() As Boolean
        Get
            If mValidCache Is Nothing Then Return False
            Return mValidCache
        End Get
    End Property

    ' dgp rev 6/26/07 Change the Experiment Location
    ' move the selected experiment into a local cache in the flowroot depot
    ' then add a checksum for validation
    Public Function CachePathExists() As Boolean

        If (Not Valid_Run) Then Return False
        If (Not UserAssigned) Then Return False
        Return System.IO.Directory.Exists(CachePath)

    End Function

    Public Function CachePathExists(ByVal user) As Boolean

        If (Not Valid_Run) Then Return False
        If (Not Valid_User(user)) Then Return False
        Return System.IO.Directory.Exists(CachePath(user))

    End Function

    Private Function CachePathExists(ByVal user As String) As Boolean

        Dim root_path As String = System.IO.Path.Combine(FlowStructure.Depot_Root, "FCSRuns")
        Dim exp_path As String = System.IO.Path.Combine(user, FirstChecksum)
        Dim full_path = System.IO.Path.Combine(root_path, exp_path)
        Return System.IO.Directory.Exists(full_path)

    End Function

    ' dgp rev 6/26/07 Change the Experiment Location
    ' move the selected experiment into a local cache in the flowroot depot
    ' then add a checksum for validatio
    Public Function CacheRemove(ByVal user As String) As Boolean

        If (Not CachePathExists(user)) Then Return True

        CacheRemove = False
        Try
            System.IO.Directory.Delete(CachePath, True)
            CacheRemove = True
            mStatus = "Cache Deleted"
        Catch ex As Exception
            mStatus = "Cache Delete Failure"
        End Try

    End Function


    ' dgp rev 6/26/07 Change the Experiment Location
    ' move the selected experiment into a local cache in the flowroot depot
    ' then add a checksum for validatio
    Private Function CacheValidate(ByVal user As String) As Boolean

        CacheValidate = False
        mStatus = "Directory Failure"

        If (CachePathExists(user)) Then
            If (AllFCSSub(CachePath)) Then Return True
        End If

    End Function

    ' dgp rev 6/1/09 Cache Path
    Public ReadOnly Property CachePath(ByVal user) As String
        Get
            If (Not Valid_Run) Then Return ""
            Dim root_path As String = System.IO.Path.Combine(FlowStructure.Depot_Root, "FCSRuns")
            Dim exp_path As String = System.IO.Path.Combine(user, FirstChecksum)
            Return System.IO.Path.Combine(root_path, exp_path)

        End Get
    End Property

    ' dgp rev 6/1/09 Cache Path
    Public ReadOnly Property CachePath() As String
        Get
            If mCachePath Is Nothing Then
                If (Valid_Run And UserAssigned) Then
                    Dim root_path As String = System.IO.Path.Combine(FlowStructure.Depot_Root, "FCSRuns")
                    Dim exp_path As String = System.IO.Path.Combine(NCIUser, FirstChecksum)
                    mCachePath = System.IO.Path.Combine(root_path, exp_path)
                End If
            End If
            Return mCachePath
        End Get
    End Property
    ' xyzzy
    ' dgp rev 10/1/09 Server Checksum
    Public ReadOnly Property ServerChkSumPath(ByVal user)
        Get
            If (Valid_Run) Then
                Return System.IO.Path.Combine(FlowStructure.ServerChksum_Root, System.IO.Path.Combine("Checksums", mNCIUser))
            End If
            Return ""
        End Get
    End Property

    ' dgp rev 10/1/09 Server Checksum
    Public ReadOnly Property ServerChkSumPath()
        Get
            If (mServerChksumPath Is Nothing) Then
                If (Valid_Run And UserAssigned) Then
                    mServerChksumPath = System.IO.Path.Combine(FlowStructure.ServerChksum_Root, System.IO.Path.Combine("Checksums", mNCIUser))
                End If
            End If
            Return mServerChksumPath
        End Get
    End Property

    ' dgp rev 10/2/09 Server Checksum File
    Public ReadOnly Property ServerChksumXML(ByVal user)
        Get
            If (Valid_Run) Then
                Return System.IO.Path.Combine(ServerChkSumPath(user), FirstChecksum + ".xml")
            End If
            Return ""
        End Get
    End Property
    ' dgp rev 10/2/09 Server Checksum File
    Public ReadOnly Property ServerChksumXML()
        Get
            If (mServerChksumXML Is Nothing) Then
                If (Valid_Run And UserAssigned) Then
                    mServerChksumXML = System.IO.Path.Combine(ServerChkSumPath, FirstChecksum + ".xml")
                End If
            End If
            Return mServerChksumXML
        End Get
    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property GlobalChksumXML()
        Get
            If (mGlobalChksumXML Is Nothing) Then
                If (Valid_Run) Then
                    mGlobalChksumXML = System.IO.Path.Combine(GlobalChksumPath, FirstChecksum + ".xml")
                End If
            End If
            Return mGlobalChksumXML
        End Get
    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumXML(ByVal user) As String
        Get
            If (Valid_Run) Then
                Return System.IO.Path.Combine(UserChksumPath(user), FirstChecksum + ".xml")
            End If
            Return ""
        End Get

    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumXML() As String
        Get
            If (mUserChksumXML Is Nothing) Then
                If (Valid_Run And UserAssigned) Then
                    mUserChksumXML = System.IO.Path.Combine(UserChksumPath, FirstChecksum + ".xml")
                End If
            End If
            Return mUserChksumXML
        End Get
    End Property

    ' dgp rev 5/28/09
    Private Function Recache() As Boolean

        If (CachePathExists()) Then
            Try
                System.IO.Directory.Delete(CachePath, True)
            Catch ex As Exception
            End Try
        End If
        Return XCacheData(NCIUser)

    End Function

    Private Function StageMask(ByVal user) As Int16

        Dim value = 0
        If (GlobalChkSumXMLExists) Then value = value + 1
        If (CachePathExists(user)) Then value = value + 2
        If (UserChksumXMLExists(user)) Then value = value + 4
        If (ServerChksumXMLExists(user)) Then value = value + 8
        Return value

    End Function

    Private Function StageMask() As Int16

        Dim value = 0
        If (GlobalChkSumXMLExists) Then value = value + 1
        If (CachePathExists()) Then value = value + 2
        If (UserChksumXMLExists) Then value = value + 4
        If (ServerChksumXMLExists) Then value = value + 8
        Return value

    End Function
    ' dgp rev 10/2/09 
    Private Function ValidStage(ByVal user) As Boolean

        Dim value = StageMask(user)
        Return (value = 0 Or value = 1 Or value = 7 Or value = 9)

    End Function

    ' dgp rev 10/2/09 
    Private Function ValidStage() As Boolean

        Dim value = StageMask()
        Return (value = 0 Or value = 1 Or value = 7 Or value = 9)

    End Function

    ' dgp rev 10/2/09 
    Public Function CurStage(ByVal user) As Stage

        Dim value = StageMask(user)

        If (value = 0) Then Return Stage.Nothing
        If (value = 1) Then Return Stage.Nothing
        If (value = 7) Then Return Stage.Ready
        If (value = 9) Then Return Stage.Uploaded
        Return Stage.StageError

    End Function

    ' dgp rev 10/2/09 
    Public Function CurStage() As Stage

        Dim value = StageMask()

        If (value = 0) Then Return Stage.Nothing
        If (value = 1) Then Return Stage.Nothing
        If (value = 7) Then Return Stage.Ready
        If (value = 9) Then Return Stage.Uploaded
        Return Stage.StageError

    End Function

    ' dgp rev 10/6/09 Has data been uploaded yet?
    Private Function ChecksumServer() As Boolean

        ChecksumServer = False

    End Function

    Private mServerChecksum

    ' dgp rev 10/6/09 clear any existing cache
    Private Function ClearCache(ByVal user) As Boolean

        If (CachePathExists(user)) Then
            Return CacheRemove(user)
        End If
        Return True

    End Function

    ' dgp rev 5/28/09 Attempt to Cache Data upon setting User
    Private Function CacheData(ByVal user As String) As Boolean

        CacheData = False
        If (Not Valid_Run) Then Return False
        If (Not Valid_User(user)) Then Return False

        mServerChecksum = ChecksumServer()

        If ClearCache(user) Then
            ' dgp rev 5/28/09 cache data
            If (CacheMoveFiles(user)) Then
                ' dgp rev 5/28/09 save checksums to an XML list file
                mValidCache = WriteUserChksum(user)
                Return mValidCache
            End If
        Else
        End If


    End Function

    ' dgp rev 5/28/09 Attempt to Cache Data upon setting User
    Private Function XCacheData(ByVal user As String) As Boolean

        XCacheData = False
        If (Not Valid_Run) Then Return False
        If (Not Valid_User(user)) Then Return False
        ' dgp rev 5/28/09 perhaps data is already cached
        If ValidStage(user) Then
            If CurStage(user) = Stage.Ready Then
                ' dgp rev 6/4/09 
                ' dgp rev 5/28/09 then simply compare data with XML cache list 
                If (Compare_Lists()) Then
                    If (CacheValidate(user)) Then Return True
                End If
            End If
        Else
            If StageMask(user) > 9 Or StageMask(user) = 3 Then CacheRemove(user)
        End If

        ' dgp rev 5/28/09 cache data
        If (CacheMoveFiles(user)) Then
            ' dgp rev 5/28/09 save checksums to an XML list file
            mValidCache = WriteUserChksum(user)
            Return mValidCache
        End If

    End Function

    Private mServerChksumPath = Nothing
    Private mServerChksumXML = Nothing
    Private mServerChksumPathExists = False
    Private mServerChksumXMLExists = False

    ' dgp rev 10/2/09 Server Checksum
    Public ReadOnly Property ServerChkSumPathExists() As Boolean
        Get
            If (ServerChkSumPath Is Nothing) Then Return False
            Return System.IO.Directory.Exists(ServerChkSumPath)
        End Get
    End Property

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
    Public Property NCIUser() As String
        Get
            Return mNCIUser
        End Get
        Set(ByVal value As String)
            If (value Is Nothing) Then Exit Property
            If (mNCIUser = value) Then Exit Property
            mNCIUser = value
            mUserAssigned = True
            XCacheData(value)
        End Set
    End Property


    Public Run_List_Path As String = "\\Nt-eib-10-6b16\FTP_root\runs"

    ' dgp rev 7/29/08 Valid User consists of Valid Experiment and User from List
    Public Function Valid_User(ByVal test As String) As Boolean

        If (UserList Is Nothing) Then Return False
        Return UserList.Contains(test)

    End Function

    ' dgp rev 7/29/08 Valid User consists of Valid Experiment and User from List
    Public Function Valid_User() As Boolean

        If mNCIUser Is Nothing Then Return False
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
        If (CurStage = Stage.Uploaded) Then Return True
        If (Successful_Upload) Then
            If (CachePathExists()) Then
                Try
                    System.IO.Directory.Delete(CachePath, True)
                    If (System.IO.Directory.Exists(GlobalChksumPath)) Then
                        System.IO.Directory.Move(GlobalChksumPath, ServerChkSumPath)
                    End If
                    CheckNClear = True
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
        StageError = -1
        [Nothing] = 0
        Ready = 2
        Uploaded = 3
        Stored = 4
    End Enum

    Public Enum StageError
        Invalid = -1
        Mismatch = -2
        CacheError = -3
        [Nothing] = 0
    End Enum

    ' dgp rev 11/3/08 Color of Current Stage
    Public ReadOnly Property StageColor() As System.Drawing.Color
        Get
            If ValidStage() Then
                Select Case CurStage()
                    Case Stage.Nothing
                        Return Drawing.Color.White
                    Case Stage.Ready
                        Return Drawing.Color.GreenYellow
                    Case Stage.Uploaded
                        Return Drawing.Color.Green
                    Case Stage.Stored
                        Return Drawing.Color.LightGreen
                End Select
            Else
                Select Case CurStage()
                    Case StageError.Invalid
                        Return Drawing.Color.Red
                    Case StageError.Mismatch
                        Return Drawing.Color.Red
                    Case StageError.CacheError
                        Return Drawing.Color.Aqua
                End Select
            End If

        End Get
    End Property

    ' dgp rev 6/22/07 checksum of first file to uniquely ident run or experiment
    Private mFirstChecksum = Nothing

    ' dgp rev 5/28/09 First Check Sum 
    Public ReadOnly Property FirstChecksum() As String
        Get
            If mFirstChecksum Is Nothing Then mFirstChecksum = Me.FCS_Files(0).ChkSumStr
            Return mFirstChecksum
        End Get
    End Property
    ' 8/6/08 Checksum Exists
    Private mGlobalChksumPath
    Private mGlobalChksumPathExists
    Private mGlobalChksumXML
    Private mGlobalChksumXMLExists
    Private mUserChksumPath
    Private mUserChksumPathExists
    Private mUserChksumXML
    Private mUserChksumXMLExists

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumXMLExists(ByVal user) As Boolean
        Get
            Return System.IO.File.Exists(UserChksumXML(user))
        End Get
    End Property
    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumXMLExists() As Boolean
        Get
            If (UserChksumXML Is Nothing) Then Return False
            Return System.IO.File.Exists(UserChksumXML)
        End Get
    End Property
    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property ServerChksumXMLExists(ByVal user) As Boolean
        Get
            Return System.IO.File.Exists(ServerChksumXML(user))
        End Get
    End Property
    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property ServerChksumXMLExists() As Boolean
        Get
            If (ServerChksumXML Is Nothing) Then Return False
            Return System.IO.File.Exists(ServerChksumXML)
        End Get
    End Property
    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumPathExists() As Boolean
        Get
            If (UserChksumPath Is Nothing) Then Return False
            Return System.IO.Directory.Exists(UserChksumPath)
        End Get
    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumPath(ByVal user)
        Get
            If (Valid_Run) Then
                Return System.IO.Path.Combine(FlowStructure.Depot_Root, System.IO.Path.Combine(user, "Checksums"))
            End If
            Return ""
        End Get
    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumPath()
        Get
            If (mUserChksumPath Is Nothing) Then
                If (Valid_Run And UserAssigned) Then
                    mUserChksumPath = System.IO.Path.Combine(FlowStructure.Depot_Root, System.IO.Path.Combine(NCIUser, "Checksums"))
                End If
            End If
            Return mUserChksumPath
        End Get
    End Property

    Public ReadOnly Property GlobalChksumPath()
        Get
            If (mGlobalChksumPath Is Nothing) Then
                mGlobalChksumPath = System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums")
            End If
            Return mGlobalChksumPath
        End Get
    End Property

    Public ReadOnly Property GlobalChkSumXMLExists() As Boolean
        Get
            If (Valid_Run()) Then
                Return System.IO.File.Exists(GlobalChksumXML)
            End If
            Return False
        End Get
    End Property

    ' dgp rev 8/6/08 Checksum Path, may or may not exist
    Public ReadOnly Property ChkSum_Root() As String
        Get
            Return System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums")
        End Get
    End Property

    ' dgp rev 8/6/08 Checksum Path, may or may not exist
    Public ReadOnly Property ChkSum_Path() As String
        Get
            If (Valid_Run()) Then Return System.IO.Path.Combine(ChkSum_Root, FirstChecksum + ".xml")
            Return ""
        End Get
    End Property

    ' dgp rev 8/6/08 Checksum Path, may or may not exist
    Public ReadOnly Property ChkSum_Path(ByVal user) As String
        Get
            If (Valid_Run()) Then
                mGlobalChksumPath = System.IO.Path.Combine(System.IO.Path.Combine(ChkSum_Root, user), FirstChecksum + ".xml")
                Return mGlobalChksumPath
            End If
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
                If (ValidStage()) Then
                    If (CurStage() = Stage.Uploaded) Then
                        Store_Away = Upload_Info.Store_Exper()
                        mStoredFlag = Store_Away
                        If (Remove_Original()) Then

                        End If
                    End If
                End If
            Case FCSType.Run
                If (ValidStage()) Then
                    If (CurStage() = Stage.Uploaded) Then
                        Store_Away = Upload_Info.Store_Files()
                        mStoredFlag = Store_Away
                    End If
                End If
        End Select

    End Function

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
        m_chksum_list = Nothing
        mServerChksumPath = Nothing
        mServerChksumXML = Nothing
        mGlobalChksumPath = Nothing
        mGlobalChksumXML = Nothing
        mUserChksumPath = Nothing
        mUserChksumXML = Nothing

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
    ' dgp rev 6/1/09 File Checksum List
    Private mFile_ChkSum_List As ArrayList = Nothing

    ' dgp rev 6/26/07 Calculate single checksum of FCS binary data block
    Public Function Calc_FCS_Chksum(ByVal file As String) As String

        If (Not System.IO.File.Exists(file)) Then Return ""

        Dim objFCS As FCS_Classes.FCS_File

        mFile_ChkSum_List = New ArrayList
        objFCS = New FCS_Classes.FCS_File(file)
        Calc_FCS_Chksum = objFCS.ChkSumStr

    End Function
    ' dgp rev 5/27/09 Calculate the checksum list only once
    Public ReadOnly Property File_ChkSum_List() As ArrayList
        Get
            If mFile_ChkSum_List Is Nothing Then Calc_ChkSm_All()
            Return mFile_ChkSum_List
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


    ' dgp rev 6/3/09 Find a Checksum file under the user sections
    Private Function FindUsers() As ArrayList

        Dim chksum = FirstChecksum

        Dim root = System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums")
        Dim usrfld
        Dim chkfil
        Dim CurUsr
        FindUsers = New ArrayList
        If Not System.IO.Directory.Exists(root) Then Exit Function
        For Each usrfld In System.IO.Directory.GetDirectories(root)
            CurUsr = System.IO.Path.GetFileName(usrfld)
            For Each chkfil In System.IO.Directory.GetFiles(usrfld)
                If (chksum = System.IO.Path.GetFileNameWithoutExtension(chkfil)) Then
                    FindUsers.Add(CurUsr)
                End If
            Next
        Next

    End Function

    ' dgp rev 6/4/09 Write Global Checksum
    Private Function WriteUserChksum(ByVal user) As Boolean

        If (Valid_Run And UserAssigned) Then
            If (Utility.Create_Tree(UserChksumPath)) Then
                Try
                    Dim XML_DataSet = CreateDataset()
                    XML_DataSet.WriteXml(UserChksumXML, XmlWriteMode.WriteSchema)
                    WriteUserChksum = True
                Catch ex As Exception

                End Try
            End If
        End If

    End Function

    ' dgp rev 6/4/09 Write Global Checksum
    Private Function WriteGlobalChksum() As Boolean

        If (Valid_Run) Then
            If (Utility.Create_Tree(GlobalChksumPath)) Then
                Try
                    Dim XML_DataSet = CreateDataset()
                    XML_DataSet.WriteXml(GlobalChksumXML, XmlWriteMode.WriteSchema)
                    WriteGlobalChksum = True
                Catch ex As Exception

                End Try
            End If
        End If

    End Function

    ' dgp rev 4/5/07 Save the Checksum in XML format
    ' dgp rev 6/25/07 Save Checksum for each file
    Public Function CreateDataset() As DataSet

        CreateDataset = New DataSet("Validator")
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
        For idx = 0 To File_ChkSum_List.Count - 1
            ' Then add the new row to the collection.
            myRow(idx) = File_ChkSum_List(idx).ToString
        Next
        XML_Table.Rows.Add(myRow)
        CreateDataset.Tables.Add(XML_Table)

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
        For idx = 0 To File_ChkSum_List.Count - 1
            ' Then add the new row to the collection.
            myRow(idx) = File_ChkSum_List(idx).ToString
        Next
        XML_Table.Rows.Add(myRow)
        XML_DataSet.Tables.Add(XML_Table)
        If (Utility.Create_Tree(DataInfoPath)) Then
            Try
                XML_DataSet.WriteXml(System.IO.Path.Combine(DataInfoPath, FirstChecksum), XmlWriteMode.WriteSchema)
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

    ' dgp rev 5/27/09 Return Status
    Public ReadOnly Property Status() As String
        Get
            Return mStatus
        End Get
    End Property

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

        Server_Path = String.Format("\\{0}{1}", Upload_Info.DataServer, Upload_Info.Upload_Root)
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

    ' dgp rev 6/26/07 Change the Experiment Location
    ' move the selected experiment into a local cache in the flowroot depot
    ' then add a checksum for validatio
    Private Function CacheMoveFiles(ByVal user As String) As Boolean

        CacheMoveFiles = False
        mStatus = "Already Cached"

        If (CachePathExists(user)) Then Return True
        If (Utility.Create_Tree(CachePath(user))) Then
            Try
                mFile_ChkSum_List = Nothing
                Move_FCS(Orig_Path, CachePath)
                CacheMoveFiles = True
                mStatus = "Successful Cache"
            Catch ex As Exception
                mStatus = "Caching Failure"
            End Try
        Else
            mStatus = "Cache Creation Failure"
        End If

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

    Private mXML_ChkSum_List = Nothing

    ' dgp rev 6/4/09 XML Checksum List
    Public ReadOnly Property XML_ChkSum_List() As ArrayList
        Get
            If mXML_ChkSum_List Is Nothing Then mXML_ChkSum_List = Load_XML_ChkSum()
            Return mXML_ChkSum_List
        End Get
    End Property

    ' dgp rev 6/26/07 
    Private Function Load_XML_ChkSum() As ArrayList

        Load_XML_ChkSum = New ArrayList

        If (Not GlobalChkSumXMLExists) Then Exit Function

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

        If (nodes.Count > 0) Then
            '            Log_Info("Node Found")
            ' check sum list read from XML file
            ' must compare to checksums from actual files
            For Each node In nodes.Item(0).ChildNodes
                Load_XML_ChkSum.Add(node.InnerText)
            Next
        End If

    End Function

    ' dgp rev 6/26/07 Compare the Actual File List with the saved XML list
    Public Function Compare_Lists() As Boolean

        Compare_Lists = False
        mStatus = "Actual Checksum Error"
        If (File_ChkSum_List.Count = 0) Then Exit Function
        mStatus = "XML Checksum Error"
        If (XML_ChkSum_List.Count = 0) Then Exit Function
        mStatus = "File Count Mismatch"
        If (XML_ChkSum_List.Count <> File_ChkSum_List.Count) Then Exit Function

        File_ChkSum_List.Sort()
        XML_ChkSum_List.Sort()

        Dim idx
        mStatus = "Checksum Mismatch"
        For idx = 0 To XML_ChkSum_List.Count - 1
            If (Not File_ChkSum_List.Contains(XML_ChkSum_List.Item(idx))) Then Exit Function
        Next

        mStatus = "Valid Checksum"
        Return True

    End Function


    ' dgp rev 6/26/07 Verify that the checksum matches the file checksum
    Private Function Verify_Checksum() As Boolean

        Verify_Checksum = False

        If (GlobalChkSumXMLExists) Then
            Return Compare_Lists()
        Else
            mStatus = "No XML checksum file"
        End If

        Return False

    End Function

    ' dgp rev 6/25/07 Prep experiment
    ' errors - mismatch, caching, or validation
    Public Function Prep() As Boolean

        If (ValidStage()) Then
            If (CurStage() = Stage.Ready) Then
                ' dgp rev 7/29/08 experiment needs to be cached and a checksum calculated
                Return Verify_Checksum()
            End If
        End If

        Return ValidStage()

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

        If (CachePathExists()) Then
            Total_Attempts = System.IO.Directory.GetFiles(CachePath).Length
            Utility.Create_Tree(Backup_Path)

            RaiseEvent evtRepeat("Backup...")
            For Each SrcFil In System.IO.Directory.GetFiles(CachePath)
                RaiseEvent evtRepeat(System.IO.Path.GetFileName(SrcFil))
                DestFil = System.IO.Path.Combine(Backup_Path, System.IO.Path.GetFileName(SrcFil))
                System.IO.File.Copy(SrcFil, DestFil, True)
                If System.IO.File.Exists(DestFil) Then Good_Backup = Good_Backup + 1
            Next
        End If

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

    ' dgp rev 6/11/09 list of checksums for all valid FCS files in the run
    Private m_chksum_list As New Hashtable
    Public ReadOnly Property ChkSum_List() As Hashtable
        Get
            If m_chksum_list Is Nothing Then Calc_ChkSum_List()
            Return m_chksum_list
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

        m_chksum_list = New Hashtable

        For Each item In FCS_List
            ' dgp rev 6/11/09 why .fcs extension, rather use Valid flag
            '            If (LCase(item.Extension) = ".fcs") Then
            HEXStr = ""
            objFCS = New FCS_File(item.FullName.ToString)
            If (objFCS.Valid) Then
                chksum = objFCS.Calc_ChkSum
                For Each bite In chksum
                    HEXStr = HEXStr + (String.Format("{0:X2}", bite))
                Next
                m_chksum_list.Add(HEXStr.ToString(), item.Name)
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

    ' dgp rev 6/26/07 Fill File Checksum List with Calculated Checksums
    Private Sub Calc_ChkSm_All()

        Dim cnt As Integer = Me.FCS_cnt
        Dim idx

        mFile_ChkSum_List = New ArrayList
        For idx = 0 To cnt - 1
            mFile_ChkSum_List.Add(FCS_Files(idx).ChkSumStr())
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
            If (System.IO.Path.GetExtension(spec).ToLower = ".xml") Then
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
                        '                        FirstRunFile = tub_node.Item("data_filename").InnerText
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
    Private mValid_Exper = Nothing
    Public ReadOnly Property Valid_Exper() As Boolean
        Get
            If mValid_Exper Is Nothing Then Validate_Exp()
            Return mValid_Exper
        End Get
    End Property

    ' dgp rev 6/19/07 Validate Experiment XML and FCS files
    ' dgp rev 12/14/07 Add a status member to record validation results
    Public Function Validate_Exp() As Boolean

        Dim tube_cnt As Integer = 0
        Dim missing As New ArrayList

        Dim btim As Date = Now()

        Validate_Exp = False

        ' scan all the FCS files in the path
        If Not Scan_Path() Then Exit Function

        Dim item
        Extract_DBFiles()
        For Each item In DBFiles
            If (mRealFiles.Contains(item)) Then
                ' remove valid files from list
                mRealFiles.RemoveAt(mRealFiles.IndexOf(item))
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
        mValid_Exper = (missing.Count = 0 And mRealFiles.Count = 0)

        If (mValid_Exper) Then
            mStatus = "XML Validated"
        Else
            If (missing.Count > 0) Then
                mStatus = "Missing XML entries"
            Else
                mStatus = "Missing FCS files"
            End If
        End If

        ' Calc_ChkSm_All()

        Return mValid_Exper

    End Function

    Public xml_cksm_doc As Xml.XmlDocument

    ' dgp rev 3/11/09 Remove CurRun, perhaps a valid backup must be confirmed
    Public Function Remove_Original() As Boolean

        ' dgp rev 7/29/08 in order to create the CurCache we need a valid experiment
        ' and a valid root
        Remove_Original = False
        If (CacheValidate(NCIUser)) Then
            Try
                System.IO.Directory.Delete(Orig_Path)
                Remove_Original = True
            Catch ex As Exception
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
