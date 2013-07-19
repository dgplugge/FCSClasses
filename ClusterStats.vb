Imports System.Xml

' Name:     Cluster Statistics Class
' Author:   Donald G Plugge
' Date:     2/10/2011
' Purpose:  Class to handle the stat_table.xml cluster statistics file

Public Class ClusterStats

    ' dgp rev 2/11/2011 
    Private Shared mClusterStatFile As String
    Public Shared Property ClusterStatFile As String
        Get
            Return mClusterStatFile
        End Get
        Set(ByVal value As String)
            mClusterStatFile = value
        End Set
    End Property

    ' dgp rev 2/14/2011
    Private Shared mMyTable As DataTable
    Public Shared ReadOnly Property MyTable As DataTable
        Get
            Return mMyTable
        End Get
    End Property

    ' Create a new table with headings
    Public Shared Sub CreateHeader()

        Dim idx As Int16
        Dim nc As DataColumn

        mMyTable = New DataTable("Cluster Statistics")
        nc = New DataColumn
        nc.ColumnName = "Cluster"
        nc.Caption = nc.ColumnName
        mMyTable.Columns.Add(nc)
        nc = New DataColumn
        nc.ColumnName = "Events"
        nc.Caption = nc.ColumnName
        mMyTable.Columns.Add(nc)

        For idx = 0 To mParNames.Count - 1
            nc = New DataColumn
            nc.ColumnName = mParNames(idx)
            nc.Caption = mParNames(idx)
            mMyTable.Columns.Add(nc)
        Next

    End Sub
    ' dgp rev 2/11/2011
    Private Shared Sub FillTable(ByVal RowInfo As ArrayList)

        Dim nr As DataRow = mMyTable.NewRow

        Dim idx

        For idx = 0 To nr.ItemArray.Length - 1
            nr.Item(idx) = RowInfo.Item(idx)
        Next
        mMyTable.Rows.Add(nr)

    End Sub

    Private Shared mClusterStats As System.Xml.XmlDataDocument

    Private Shared mParNames As ArrayList
    Private Shared mMeansTable As ArrayList

    ' dgp rev 2/11/2011
    Private Shared Function ExtractMeans(ByVal mXMLDoc As System.Xml.XmlDataDocument) As Boolean

        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        mMeansTable = New ArrayList
        Dim row As ArrayList
        Dim ParIdx As Integer
        Try
            'Get the list of name nodes 
            m_nodelist = mXMLDoc.SelectNodes("/Environment/FlowControl/Statistics/ClusterMeans/Means")
            'Loop through the nodes
            For Each m_node In m_nodelist
                row = New ArrayList
                row.Add(m_node.Attributes.Item(0).Value)
                row.Add(m_node.Attributes.Item(1).Value)
                For ParIdx = 2 To m_node.Attributes.Count - 1
                    'Get the Gender Attribute Value
                    row.Add(m_node.Attributes.Item(ParIdx).Value)
                    'Dim genderAttribute = m_node.Attributes.GetNamedItem("").Value
                    'Get the firstName Element Value
                    'Dim firstNameValue = m_node.ChildNodes.Item(0).InnerText
                    'Get the lastName Element Value
                    'Dim lastNameValue = m_node.ChildNodes.Item(1).InnerText
                    'Write Result to the Console
                Next
                FillTable(row)
            Next
        Catch ex As Exception

        End Try
        Return mParNames.Count > 0

    End Function

    ' dgp rev 2/15/2011
    Private Shared mClusterInfo As ArrayList
    Public Shared ReadOnly Property ClusterInfo As ArrayList
        Get
            Return mClusterInfo
        End Get
    End Property

    ' dgp rev 2/14/2011
    Private Shared Function ExtractName(ByVal mXMLDoc As System.Xml.XmlDataDocument) As Boolean

        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        mParNames = New ArrayList
        mClusterInfo = New ArrayList
        Try
            'Get the list of name nodes 
            m_nodelist = mXMLDoc.SelectNodes("/Environment/FlowControl/Statistics")
            'Loop through the nodes
            For Each m_node In m_nodelist.Item(0).Attributes
                'Get the Unique Attribute Value
                mClusterInfo.Add(m_node.Value)
            Next
        Catch ex As Exception
            Return True
        End Try
        Return True

    End Function

    ' dgp rev 2/14/2011
    Private Shared Function ExtractParameters(ByVal mXMLDoc As System.Xml.XmlDataDocument) As Boolean

        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        mParNames = New ArrayList
        Try
            'Get the list of name nodes 
            m_nodelist = mXMLDoc.SelectNodes("/Environment/FlowControl/Statistics/Parameters")
            'Loop through the nodes
            For Each m_node In m_nodelist.Item(0).Attributes
                'Get the Gender Attribute Value
                mParNames.Add(m_node.Value)
                'Dim genderAttribute = m_node.Attributes.GetNamedItem("").Value
                'Get the firstName Element Value
                'Dim firstNameValue = m_node.ChildNodes.Item(0).InnerText
                'Get the lastName Element Value
                'Dim lastNameValue = m_node.ChildNodes.Item(1).InnerText
                'Write Result to the Console
            Next
        Catch ex As Exception

        End Try
        Return mParNames.Count > 0

    End Function

    ' dgp rev 2/14/2011
    Private Sub HoldTable()

        Dim ds As DataSet = New DataSet("Cluster Statistics")
        Dim info = ds.ReadXml(ClusterStatFile)
        Dim dt As DataTable
        For Each dt In ds.Tables
            Dim cnt = dt.Columns.Count
        Next

    End Sub

    ' dgp rev 2/10/2011 Process the new cluster stats file
    Public Shared Function ProcessClusters() As Boolean

        If System.IO.File.Exists(ClusterStatFile) Then

            Dim mXMLDoc As New System.Xml.XmlDataDocument()
            Try
                mXMLDoc.Load(ClusterStatFile)
                If (ExtractName(mXMLDoc)) Then
                    If (ExtractParameters(mXMLDoc)) Then
                        CreateHeader()
                        If (ExtractMeans(mXMLDoc)) Then
                        End If
                    End If
                End If
            Catch ex As Exception
                Return False

            End Try

        End If

        Return True

    End Function

End Class
