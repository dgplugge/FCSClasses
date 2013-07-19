' Name:     FCS Table Class
' Author:   Donald G Plugge
' Date:     3/17/06 
' Purpose:  Class to handle customized parameter anotations to incorporate
'           into the text header
' Initialize from XML, PRO or FCS Run
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Specialized
Imports HelperClasses

Public Class FCSTable

    ' The Table object is mainly a DataTable type with enhancements
    Public myTable As DataTable

    Private XML_flg As Boolean = False
    Private PRO_flg As Boolean = False

    ' internal and environment variables defined
    Private _pro_file_arr As ArrayList
    Private _pro_matrix As ArrayList
    Private mDate As String = DateTime.Now.ToString("yyMMdd")
    Public ftp_pro_file As String

    'dgp rev 8/6/2010 Work on Merge
    Private mProKeys As Dictionary(Of String, String) = Nothing
    Public ReadOnly Property ProKeys As Dictionary(Of String, String)
        Get
            If mProKeys Is Nothing Then ProtocolMatches()
            Return mProKeys
        End Get
    End Property

    ' dgp rev 8/5/2010 
    Private Sub ProtocolMatches()

        mProKeys = New Dictionary(Of String, String)
        If (Get_Keys.Count = 0) Then Return

        Dim key
        Dim item
        ' dgp rev 8/5/2010 loop thru current protocol and look for matches
        For Each item In Get_Keys()
            key = FCSAntibodies.FindMatch(item)
            If key <> "" And Not mProKeys.ContainsKey(key) Then mProKeys.Add(key, item)
        Next

    End Sub

    ' dgp rev 7/7/09 Global User
    Private Shared mGlobalUser = Environment.UserName.ToString
    Public Shared ReadOnly Property GlobalUser() As String
        Get
            Return mGlobalUser
        End Get
    End Property

    ' dgp rev 4/22/09 default to current user, however this may be changed
    Private Shared mUser = Nothing
    Public Shared Property User() As String
        Get
            If mUser Is Nothing Then mUser = GlobalUser
            Return mUser
        End Get
        Set(ByVal value As String)
            mUser = value
        End Set
    End Property

    ' dgp rev 8/10/2010 Full path 
    Private mFullpath As String = ""
    ' dgp rev 8/10/2010 Table Name with extension
    Private mFullName As String = Nothing

    Public ReadOnly Property FullSpec As String
        Get
            Return System.IO.Path.Combine(FullPath, mFullName)
        End Get
    End Property

    ' dgp rev 8/24/2011 Server path specification
    Public ReadOnly Property ServerSpec As String
        Get
            Return System.IO.Path.Combine(ServerPath, mFullName)
        End Get
    End Property

    ' dgp rev 8/10/2010 does the protocol currently exist
    Private mExists As Boolean = False
    Public ReadOnly Property Exists As Boolean
        Get
            Return mExists
        End Get
    End Property

    ' dgp rev 8/`10/2010 Full path name is set via SetProtocolSpec
    Public ReadOnly Property FullPath As String
        Get
            Return mFullpath
        End Get
    End Property

    ' dgp rev 8/10/2010 Set full specification of protocol file
    Public WriteOnly Property SetProtocolSpec() As String
        Set(ByVal value As String)
            mExists = System.IO.File.Exists(value)
            If (mExists) Then
                mFullpath = System.IO.Path.GetDirectoryName(value)
                mFullName = System.IO.Path.GetFileName(value)
            Else
                If (System.IO.Directory.Exists(value)) Then
                    mFullpath = System.IO.Path.GetFullPath(value)
                End If
            End If
        End Set
    End Property


    ' dgp rev 8/10/2010 Table Name without extension
    Public Property ProtocolName() As String
        Get
            If mFullName Is Nothing Then mFullName = User + "_" + mDate + ".xml"
            Return System.IO.Path.GetFileNameWithoutExtension(mFullName)
        End Get
        Set(ByVal value As String)
            mFullName = System.IO.Path.GetFileNameWithoutExtension(value) + ".xml"
        End Set
    End Property

    ' dgp rev 7/15/09 Path to Server
    Private mServerPath = Nothing
    Public ReadOnly Property ServerPath() As String
        Get
            mServerPath = String.Format("\\{0}\{1}\Users\{2}", FlowServer.FlowServer, FlowServer.ShareFlow, User)
            mServerPath = System.IO.Path.Combine(mServerPath, "Settings")
            mServerPath = System.IO.Path.Combine(mServerPath, "Tables")
            Return mServerPath
        End Get
    End Property


    ' Matrix of Table Info
    Private mMatrix As ArrayList
    Public Property Matrix() As ArrayList
        Get
            Return mMatrix
        End Get
        Set(ByVal value As ArrayList)
            mMatrix = value
        End Set
    End Property

    Private mNewLocation As Boolean = False
    Public Property NewLocation() As Boolean
        Get
            Return mNewLocation
        End Get
        Set(ByVal value As Boolean)
            mNewLocation = value
        End Set
    End Property

    Private mTargetPath As String
    Public Property TargetPath() As String
        Get
            Return mTargetPath
        End Get
        Set(ByVal value As String)
            mTargetPath = value
        End Set
    End Property

    Private mRenameFlg As Boolean
    Public Property RenameFlag() As Boolean
        Get
            Return mRenameFlg
        End Get
        Set(ByVal value As Boolean)
            mRenameFlg = value
        End Set
    End Property

    ' dgp rev 3/23/06 Merge the keys and values into the current file object
    Private Function File_Merge(ByVal objFile As FCS_File, ByVal keys As ArrayList, ByVal vals As ArrayList) As Boolean

        Dim idx As Int16
        Dim key_list As String = ""
        Dim item As String
        Dim marker As String = "EIB Table"

        For Each item In keys
            key_list += item + ","
        Next

        ' add each table item
        For idx = 0 To keys.Count - 1
            If (objFile.Header.ContainsKey(keys(idx))) Then
                objFile.Header.Item(keys(idx)) = vals(idx)
            Else
                objFile.Header.Add(keys(idx), vals(idx))
            End If
        Next
        ' add a marker to list the EIB Table
        If (objFile.Header.ContainsKey(marker)) Then
            objFile.Header.Item(marker) = key_list
        Else
            objFile.Header.Add(marker, key_list)
        End If

        Return True

    End Function


    ' dgp rev 3/23/06 Extract the key values from the current file object
    Public Function File_Extract(ByVal objFile As FCS_File) As ArrayList

        Dim item As String
        Dim Table_Info As New ArrayList

        For Each item In objFile.Header.Keys
            If (item(0) = "#") Then
                Table_Info.Add(objFile.Header.Item(item))
            End If
            If (item = "$SMNO") Then
                Table_Info.Add(objFile.Header.Item(item))
            End If
        Next

        Return Table_Info

    End Function
    ' dgp rev 7/24/07 the table is valid if it has rows and columns defined
    Private mValid As Boolean
    Public ReadOnly Property Valid() As Boolean
        Get
            If (myTable Is Nothing) Then Return False
            '            Return (myTable.Rows.Count > 0 And myTable.Columns.Count > 0)
            Return (myTable.Columns.Count > 0)
        End Get
    End Property
    ' .PRO comment block
    Private m_comments As ArrayList
    Public Property Comments() As ArrayList
        Get
            If (m_comments Is Nothing) Then
                m_comments = New ArrayList
                m_comments.Add("! Protocol file name: ")
                m_comments.Add("! This protocol is a new format XML file.")
            End If
            Return m_comments
        End Get
        Set(ByVal Value As ArrayList)
            m_comments = Value
        End Set
    End Property

    ' .PRO BEGIN/END block
    ' dgp rev 11/17/08 give the Begin/End Block default values if empty
    Private m_beblock As ArrayList
    Public Property BEBlock() As ArrayList
        Get
            If (m_beblock Is Nothing) Then
                m_beblock = New ArrayList
                m_beblock.Add(".BEGIN")
                m_beblock.Add("PROT$FIL=" + Me.ProtocolName)
                m_beblock.Add(".END")
            End If
            Return m_beblock
        End Get
        Set(ByVal Value As ArrayList)
            m_beblock = Value
        End Set
    End Property

    ' table array contain the old style .PRO table
    Private table_lst As New ArrayList
    ' Get all the keys for the full table
    Public Function Get_Keys() As ArrayList

        Dim tmp As New ArrayList
        Dim col As DataColumn

        If (Not myTable Is Nothing) Then
            For Each col In myTable.Columns
                tmp.Add(col.ColumnName.ToString)
            Next
        End If
        Return tmp

    End Function

    ' Get all the keys for the full table
    Public Function Get_Values(ByVal idx As Int16) As ArrayList

        Dim item
        Dim tmp As New ArrayList

        For Each item In myTable.Rows(idx).ItemArray
            tmp.Add(item)
        Next
        Return tmp

    End Function

    ' dgp rev 4/12/06 Create the obsolete PRO style table
    Public Sub Table_to_Pro()

        Dim keys As ArrayList

        keys = Get_Keys()
        Dim item
        Dim tmp As New ArrayList
        Dim smno_idx As Int16
        Dim key_idx As Int16 = 0

        tmp.AddRange(Comments)
        If (Not BEBlock Is Nothing) Then tmp.AddRange(BEBlock)
        For smno_idx = 1 To myTable.Rows.Count
            tmp.Add(".PAUSE Load Sample #" + Format(smno_idx, "000"))
            key_idx = 0
            For Each item In myTable.Rows(smno_idx - 1).ItemArray
                tmp.Add(keys(key_idx) + "=" + item)
                key_idx += 1
            Next
            tmp.Add(".ACQ")
        Next
        _pro_file_arr = tmp

    End Sub
    ' create the obsolete .PRO file given a file name
    Private Function Create_Pro(ByVal file_name As String) As Boolean

        Dim item As String

        Dim ts As New StreamWriter(file_name, False)
        For Each item In _pro_file_arr
            ts.WriteLine(item)
        Next
        ts.Close()
        Create_Pro = True

    End Function
    ' create the obsolete .PRO file from existing table path and name
    Public Function Create_Pro() As Boolean

        ftp_pro_file = System.IO.Path.Combine(Me.FullPath, "x_" + Me.ProtocolName + ".pro")
        Dim item As String

        Dim ts As New StreamWriter(ftp_pro_file, False)
        For Each item In _pro_file_arr
            ts.WriteLine(item)
        Next
        ts.Close()
        Create_Pro = True

    End Function

    ' dgp rev 4/15/09 Create new table from DataTable 
    Public Sub New(ByVal DT As DataTable)

        myTable = DT

    End Sub

    ' Create a new table with headings
    Public Sub New(ByVal Col_arr As ArrayList)

        Dim idx As Int16
        Dim nc As DataColumn

        myTable = New DataTable("New Table")
        nc = New DataColumn
        nc.ColumnName = "$SMNO"
        nc.Caption = nc.ColumnName
        myTable.Columns.Add(nc)

        For idx = 0 To Col_arr.Count - 1
            If (Not Col_arr(idx) = "$SMNO") Then
                nc = New DataColumn
                nc.ColumnName = Col_arr(idx)
                nc.Caption = Col_arr(idx)
                myTable.Columns.Add(nc)
            End If
        Next

        Dim nr As DataRow = myTable.NewRow
        myTable.Rows.Add(nr)

        '        Dim myRow As DataRow
        ' dgp rev 11/30/07 force a single row for new tables
        Tubes = 1

        '       For idx = 0 To Tubes - 1
        '        myRow = myTable.NewRow()
        '       myRow(0) = Format(idx + 1, "000")
        '      myTable.Rows.Add(myRow)
        '     Next

    End Sub



    ' data set to hold the current table information
    ' dataset is fill from .PRO file or from .XML file
    ' take table_lst and place it into the dataset
    Public Function Pro_Matrix_to_Table() As Boolean

        Dim heading As New ArrayList
        Dim idx As Int16

        myTable = New DataTable("PRO Table")

        ' create each column
        heading = _pro_matrix(0)
        For idx = 0 To heading.Count - 1
            Dim nc As New DataColumn
            If (heading.Item(idx) = "") Then heading.Item(idx) = "#Column" + CStr(idx + 1)
            nc.ColumnName = heading.Item(idx)
            nc.Caption = heading.Item(idx)
            myTable.Columns.Add(nc)
        Next

        Dim file_idx, key_idx As Int16
        Dim myRow As DataRow

        ' fill in the rows
        For file_idx = 1 To _pro_matrix.Count - 1
            ' Once a table has been created, use the NewRow to create a DataRow.
            myRow = myTable.NewRow()
            heading = _pro_matrix(file_idx)
            For key_idx = 0 To heading.Count - 1
                ' Then add the new row to the collection.
                myRow(key_idx) = heading(key_idx)
            Next
            myTable.Rows.Add(myRow)
        Next
        ParCnt = myTable.Columns.Count
        Tubes = myTable.Rows.Count

    End Function

    ' parse the old style .PRO table
    Private Function Pro_to_Matrix() As Boolean

        Dim rec_idx As Int16 = 0
        Dim tmp_arr As New ArrayList

        Pro_to_Matrix = False
        ' loop thru the PRO table and enter into array
        ' extract the comments
        While (_pro_file_arr.Item(rec_idx).Chars(0) = "!" And rec_idx < _pro_file_arr.Count())
            tmp_arr.Add(_pro_file_arr.Item(rec_idx).ToString)
            rec_idx += 1
        End While
        Comments = tmp_arr

        ' extract BEGIN/END block
        Dim be_arr As New ArrayList
        While (_pro_file_arr.Item(rec_idx).ToLower.IndexOf(".pause") < 0 And rec_idx < _pro_file_arr.Count())
            If (_pro_file_arr.Item(rec_idx).ToLower.IndexOf("prot$fil") >= 0) Then
                Me.ProtocolName = _pro_file_arr.Item(rec_idx).Substring(_pro_file_arr.Item(rec_idx).IndexOf("=") + 1).Trim(" ")
            End If
            be_arr.Add(_pro_file_arr.Item(rec_idx).ToString)
            rec_idx += 1
        End While
        BEBlock = be_arr

        Dim key_arr As New ArrayList
        Dim val_arr As New ArrayList
        rec_idx += 1 ' skip pause
        ' first set of keys and values - loop until next .pause
        While (_pro_file_arr.Item(rec_idx).ToLower.IndexOf(".acq") < 0)
            key_arr.Add(_pro_file_arr.Item(rec_idx).Substring(0, _pro_file_arr.Item(rec_idx).IndexOf("=")))
            val_arr.Add(_pro_file_arr.Item(rec_idx).Substring(_pro_file_arr.Item(rec_idx).IndexOf("=") + 1))
            rec_idx += 1
        End While
        _pro_matrix = New ArrayList
        _pro_matrix.Add(key_arr.Clone)
        _pro_matrix.Add(val_arr.Clone)
        val_arr.Clear()

        rec_idx += 2 ' skip acq and pause
        ' loop thru the PRO table and enter into array
        While (rec_idx < _pro_file_arr.Count() - 1)
            While (_pro_file_arr.Item(rec_idx).ToLower.IndexOf(".acq") < 0 And rec_idx < _pro_file_arr.Count() - 1)
                val_arr.Add(_pro_file_arr.Item(rec_idx).Substring(_pro_file_arr.Item(rec_idx).IndexOf("=") + 1))
                rec_idx += 1
            End While
            _pro_matrix.Add(val_arr.Clone)
            val_arr.Clear()
            rec_idx += 2
        End While

        Console.WriteLine("Done")
        Pro_to_Matrix = True

    End Function

    ' dgp rev 3/16/06 read the new style .XML tables
    Public Function Read_XML(ByVal filespec As String) As Boolean

        Dim ds As New DataSet

        myTable = New DataTable("XML Table")
        Read_XML = ds.ReadXml(filespec)
        If (Read_XML) Then
            myTable = ds.Tables.Item(0)
            ParCnt = myTable.Columns.Count
            Tubes = myTable.Rows.Count
            If (ds.ExtendedProperties.ContainsValue("Comments")) Then
                Comments = ds.ExtendedProperties.Item("Comments")
            End If
        End If

    End Function

    ' dgp rev 3/16/06 read the old style .PRO tables
    Public Function Read_PRO(ByVal filespec As String) As Boolean

        Dim sr As New System.IO.StreamReader(filespec)
        Dim rec_line As String

        rec_line = sr.ReadLine
        _pro_file_arr = New ArrayList
        While (Not rec_line Is Nothing)
            _pro_file_arr.Add(rec_line)
            rec_line = sr.ReadLine
        End While
        sr.Close()

        If (Pro_to_Matrix()) Then
            Read_PRO = True
        Else
            Read_PRO = False
        End If

    End Function

    ' Tube count
    Private m_tubes As Int16 = 0
    Public Property Tubes() As Int16
        Get
            Return m_tubes
        End Get
        Set(ByVal Value As Int16)
            m_tubes = Value
        End Set
    End Property

    ' parameter count
    Private m_parcnt As Int16 = 0
    Public Property ParCnt() As Int16
        Get
            Return m_parcnt
        End Get
        Set(ByVal Value As Int16)
            m_parcnt = Value
        End Set
    End Property

    ' dgp rev 7/26/07 Save the Protocol
    Public Sub Save_Table()

        Dim new_table As DataTable

        '        filepath = Path.Combine(me.Path, me.Me.ProtocolName + ".xml")
        If (Not Utility.Create_Tree(Me.ServerPath)) Then MsgBox("Directory creation error - " + ServerPath.ToString, MsgBoxStyle.Information)

        Try
            FlowServer.Impersonate.ImpersonateStart()
            If (System.IO.File.Exists(Me.ServerSpec)) Then
                If (MsgBox("Overwrite " + ProtocolName, MsgBoxStyle.OkCancel) <> MsgBoxResult.Ok) Then Exit Sub
            End If
        Catch ex As Exception
            FlowServer.Impersonate.ImpersonateStart()
            MsgBox("No access to " + Me.ProtocolName, MsgBoxStyle.OkCancel)
            Exit Sub
        End Try

        Dim SaveSet As New DataSet
        SaveSet.ExtendedProperties.Add("Comments", Me.Comments)

        Me.myTable.AcceptChanges()
        new_table = Me.myTable.Copy()
        SaveSet.Tables.Add(new_table)

        '        If (Not ObjImpersonate Is Nothing) Then ObjImpersonate.ImpersonateStart()
        Try
            SaveSet.WriteXml(Me.ServerSpec, XmlWriteMode.WriteSchema)
        Catch ex As Exception
            MsgBox("Save Error - " + Me.ProtocolName, MsgBoxStyle.Information)
            FlowServer.Impersonate.ImpersonateStop()
            Exit Sub
        End Try
        MsgBox("Protocol Saved - " + ProtocolName, MsgBoxStyle.Information)
        FlowServer.Impersonate.ImpersonateStop()

    End Sub

    Public ReadOnly Property IsXML As Boolean
        Get
            Return System.IO.Path.GetExtension(mFullName).ToLower = ".xml"
        End Get
    End Property

    ' dgp rev 7/24/07 Create a new table object from XML or PRO file
    Public Sub New(ByVal filespec As String)

        SetProtocolSpec = filespec

        If (Exists) Then
            If (IsXML) Then
                If (Read_XML(filespec)) Then
                    XML_flg = True
                End If
            ElseIf (System.IO.Path.GetExtension(filespec).ToLower = ".pro") Then
                If (Read_PRO(filespec)) Then
                    Pro_Matrix_to_Table()
                    PRO_flg = True
                End If
            End If
        End If

    End Sub

    ' dgp rev 4/22/09
    Public Sub New()

        myTable = New DataTable

    End Sub

    ' dgp rev 7/19/07 Initial instance of FCSMerge
    Public Sub New(ByVal objRun As FCSRun)

        If (objRun.Valid_Run) Then
            objRun.ExtractRunProtocol()
            '            Path = objRun.Data_Path
            Me.myTable = objRun.Protocol.myTable
            objRun = Nothing
        End If

    End Sub

    ' dgp rev 7/26/07 Save the Protocol
    Public Sub Save_Table(ByVal filename As String)

        Dim filepath As String
        Dim new_table As DataTable

        filename.ToLower.Replace(".xml", "")
        filename = filename + ".XML"

        '        filepath = Path.Combine(objFCSTable.Path, objFCSTable.filename + ".xml")
        filepath = Path.Combine(Me.ServerSpec, filename)

        Try
            FlowServer.Impersonate.ImpersonateStart()
            If (System.IO.Directory.Exists(filepath)) Then
                If (MsgBox("Overwrite " + filename, MsgBoxStyle.OkCancel) <> MsgBoxResult.Ok) Then Exit Sub
            End If
        Catch ex As Exception
            FlowServer.Impersonate.ImpersonateStart()
            MsgBox("No access to " + filename, MsgBoxStyle.OkCancel)
            Exit Sub
        End Try

        Dim SaveSet As New DataSet
        SaveSet.ExtendedProperties.Add("Comments", Comments)

        myTable.AcceptChanges()
        new_table = myTable.Copy()
        SaveSet.Tables.Add(new_table)

        '        If (Not ObjImpersonate Is Nothing) Then ObjImpersonate.ImpersonateStart()
        Try
            SaveSet.WriteXml(filepath, XmlWriteMode.WriteSchema)
        Catch ex As Exception
            MsgBox("Save Error - " + filename, MsgBoxStyle.Information)
            FlowServer.Impersonate.ImpersonateStop()
            Exit Sub
        End Try
        MsgBox("Protocol Saved - " + filename, MsgBoxStyle.Information)
        FlowServer.Impersonate.ImpersonateStop()

    End Sub

End Class
