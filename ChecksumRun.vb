' Name:     FCS Run Class
' Author:   Donald G Plugge
' Date:     3/22/06 
' Purpose:  Class to track information concerning a given FCS run
Imports System.IO
Imports System.Xml
Imports System.Security.Cryptography
Imports HelperClasses

Public Class ChecksumRun


    Private mXML_ChkSum_List = Nothing

    ' dgp rev 6/4/09 XML Checksum List
    Public ReadOnly Property XML_ChkSum_List() As ArrayList
        Get
            If mXML_ChkSum_List Is Nothing Then mXML_ChkSum_List = Load_XML_ChkSum()
            Return mXML_ChkSum_List
        End Get
    End Property

    Private xml_doc As System.Xml.XmlDocument

    ' dgp rev 6/26/07 
    Private Function Load_XML_ChkSum() As ArrayList

        Load_XML_ChkSum = New ArrayList

        If (Not GlobalChkSumXMLExists) Then Exit Function

        '        Log_Info("Loading " + ChkSum_Path)
        xml_doc = New System.Xml.XmlDocument
        xml_doc.Load(ChkSum_Path)

        Dim nodes As System.Xml.XmlNodeList
        Dim node As System.Xml.XmlNode

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

    Private mStatus As String
    Public ReadOnly Property Status As String
        Get
            Return mStatus
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

    ' dgp rev 6/26/07 Fill File Checksum List with Calculated Checksums
    Private Sub Calc_ChkSm_All()

        Dim cnt As Integer = BaseRun.FCS_cnt
        Dim idx

        mFile_ChkSum_List = New ArrayList
        For idx = 0 To cnt - 1
            mFile_ChkSum_List.Add(BaseRun.FCS_Files(idx).ChkSumStr())
        Next

    End Sub

    Private mDataInfoRoot

    ' dgp rev 4/29/09 Data Inforamation - checksum list, user, run
    Private Function DataInfoRoot()

        If (mDataInfoRoot Is Nothing) Then
            mDataInfoRoot = System.IO.Path.Combine(FlowStructure.RemoteRoot, "Settings")
            mDataInfoRoot = System.IO.Path.Combine(mDataInfoRoot, "Data")
        End If
        Return mDataInfoRoot

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

        For Each item In BaseRun.FCS_List
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


    ' dgp rev 4/5/07 Save the Checksum in XML format
    ' dgp rev 6/25/07 Save Checksum for each file
    Public Function CreateDataset() As DataSet

        CreateDataset = New DataSet("Validator")
        Dim XML_Table As New DataTable("List")

        Dim idx As Int16
        Dim nc As DataColumn
        Dim item

        For Each item In BaseRun.FCS_List
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

        If (BaseRun.Valid_Run And BaseRun.UserAssigned) Then
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

        If (BaseRun.Valid_Run) Then
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


    Private Function StageMask(ByVal user) As Int16

        Dim value = 0
        If (GlobalChkSumXMLExists) Then value = value + 1
        '        If (CachePathExists(user)) Then value = value + 2
        If (UserChksumXMLExists(user)) Then value = value + 4
        If (ServerChksumXMLExists(user)) Then value = value + 8
        Return value

    End Function

    ' dgp rev 10/2/09 
    Private Function ValidStage(ByVal user) As Boolean

        Dim value = StageMask(user)
        Return (value = 0 Or value = 1 Or value = 7 Or value = 9)

    End Function

    ' dgp rev 6/22/07 checksum of first file to uniquely ident run or experiment
    Private mFirstChecksum = Nothing

    ' dgp rev 5/28/09 First Check Sum 
    Public ReadOnly Property FirstChecksum() As String
        Get
            If mFirstChecksum Is Nothing Then mFirstChecksum = BaseRun.FCS_Files(0).ChkSumStr
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
            If (BaseRun.Valid_Run) Then
                Return System.IO.Path.Combine(FlowStructure.Depot_Root, System.IO.Path.Combine(user, "Checksums"))
            End If
            Return ""
        End Get
    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumPath()
        Get
            If (mUserChksumPath Is Nothing) Then
                If (BaseRun.Valid_Run And BaseRun.UserAssigned) Then
                    mUserChksumPath = System.IO.Path.Combine(FlowStructure.Depot_Root, System.IO.Path.Combine(BaseRun.NCIUser, "Checksums"))
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
            If (BaseRun.Valid_Run()) Then
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
            If (BaseRun.Valid_Run()) Then Return System.IO.Path.Combine(ChkSum_Root, FirstChecksum + ".xml")
            Return ""
        End Get
    End Property

    ' dgp rev 8/6/08 Checksum Path, may or may not exist
    Public ReadOnly Property ChkSum_Path(ByVal user) As String
        Get
            If (BaseRun.Valid_Run()) Then
                mGlobalChksumPath = System.IO.Path.Combine(System.IO.Path.Combine(ChkSum_Root, user), FirstChecksum + ".xml")
                Return mGlobalChksumPath
            End If
            Return ""
        End Get
    End Property

    Private mChksmRoot
    Public ReadOnly Property ChksmRoot() As String
        Get
            If (mChksmRoot Is Nothing) Then mChksmRoot = System.IO.Path.Combine(FlowStructure.Depot_Root, "Checksums")
            Return mChksmRoot
        End Get

    End Property

    ' dgp rev 10/6/09 Has data been uploaded yet?
    Private Function ChecksumServer() As Boolean

        ChecksumServer = False

    End Function

    Private mServerChecksum

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

    Private Function ChksmPathExists(ByVal user As String) As Boolean

        Dim full_path = System.IO.Path.Combine(ChksmRoot, user)
        Return System.IO.Directory.Exists(full_path)

    End Function

    ' dgp rev 6/1/09 Cache Path
    Public ReadOnly Property ChksmPath(ByVal user) As String
        Get
            If (Not BaseRun.Valid_Run) Then Return ""
            If (Not BaseRun.Valid_User(user)) Then Return ""
            Return System.IO.Path.Combine(BaseRun.Orig_Path, user)

        End Get
    End Property

    ' dgp rev 10/1/09 Server Checksum
    Public ReadOnly Property ServerChkSumPath(ByVal user)
        Get
            If (BaseRun.Valid_Run) Then
                Return System.IO.Path.Combine(FlowStructure.ServerChksum_Root, System.IO.Path.Combine("Checksums", BaseRun.NCIUser))
            End If
            Return ""
        End Get
    End Property

    ' dgp rev 10/1/09 Server Checksum
    Public ReadOnly Property ServerChkSumPath()
        Get
            If (mServerChksumPath Is Nothing) Then
                If (BaseRun.Valid_Run And BaseRun.UserAssigned) Then
                    mServerChksumPath = System.IO.Path.Combine(FlowStructure.ServerChksum_Root, System.IO.Path.Combine("Checksums", BaseRun.NCIUser))
                End If
            End If
            Return mServerChksumPath
        End Get
    End Property

    ' dgp rev 10/2/09 Server Checksum File
    Public ReadOnly Property ServerChksumXML(ByVal user)
        Get
            If (BaseRun.Valid_Run) Then
                Return System.IO.Path.Combine(ServerChkSumPath(user), FirstChecksum + ".xml")
            End If
            Return ""
        End Get
    End Property
    ' dgp rev 10/2/09 Server Checksum File
    Public ReadOnly Property ServerChksumXML()
        Get
            If (mServerChksumXML Is Nothing) Then
                If (BaseRun.Valid_Run And BaseRun.UserAssigned) Then
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
                If (BaseRun.Valid_Run) Then
                    mGlobalChksumXML = System.IO.Path.Combine(GlobalChksumPath, FirstChecksum + ".xml")
                End If
            End If
            Return mGlobalChksumXML
        End Get
    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumXML(ByVal user) As String
        Get
            If (BaseRun.Valid_Run) Then
                Return System.IO.Path.Combine(UserChksumPath(user), FirstChecksum + ".xml")
            End If
            Return ""
        End Get

    End Property

    ' dgp rev 10/2/09 User Checksum Path
    Public ReadOnly Property UserChksumXML() As String
        Get
            If (mUserChksumXML Is Nothing) Then
                If (BaseRun.Valid_Run And BaseRun.UserAssigned) Then
                    mUserChksumXML = System.IO.Path.Combine(UserChksumPath, FirstChecksum + ".xml")
                End If
            End If
            Return mUserChksumXML
        End Get
    End Property


    Private mBaseRun As FCSRun
    Public ReadOnly Property BaseRun As FCSRun
        Get
            Return mBaseRun
        End Get
    End Property

    Public Sub New(ByVal BaseRun As FCSRun)

        mServerChksumPath = Nothing
        mServerChksumXML = Nothing
        mGlobalChksumPath = Nothing
        mGlobalChksumXML = Nothing
        mUserChksumPath = Nothing
        mUserChksumXML = Nothing
        mBaseRun = BaseRun
        BaseRun = Nothing

    End Sub

End Class
