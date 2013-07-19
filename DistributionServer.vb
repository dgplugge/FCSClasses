' Author: Donald G Plugge
' Title:
' Date: 12/13/2010
' Purpose:
Imports FCS_Classes
Imports System.Xml.Linq
Imports HelperClasses

Public Class DistributionServer

    Private Shared mServerUp As Boolean
    Private Shared mPingOnce As Boolean = False

    ' dgp rev 12/15/2010 
    Private Shared mServer As String = FlowServer.FlowServer
    Public Shared ReadOnly Property Server As String
        Get
            Return mServer
        End Get
    End Property

    ' dgp rev 12/15/2010 
    Public Shared Function DistributionList() As ArrayList

        DistributionList = New ArrayList
        Dim UNC_Path As String = "\\" + Server + "\" + ShareDistribution + "\Versions"
        If (System.IO.Directory.Exists(UNC_Path)) Then
            Dim item
            For Each item In System.IO.Directory.GetDirectories(UNC_Path)
                DistributionList.Add(item.name)
            Next
        End If

    End Function

    ' dgp rev 12/15/2010 
    Public Shared ReadOnly Property ServerUp As Boolean
        Get
            If Not mPingOnce Then
                Try
                    mPingOnce = True
                    mServerUp = My.Computer.Network.Ping(Server, 1000)
                Catch ex As Exception
                    mServerUp = False
                End Try
            End If
            Return mServerUp
        End Get
    End Property

    ' dgp rev 12/15/2010 
    Public Shared ReadOnly Property PathValid As Boolean
        Get
            If Not ServerUp Then Return False
            Return True
        End Get
    End Property

    ' dgp rev 12/15/2010 
    Public Shared ReadOnly Property DistributionRoot As String
        Get
            Return String.Format("\\{0}\{1}\Versions", FlowServer.FlowServer, ShareDistribution)
        End Get
    End Property

    ' dgp rev 12/15/2010 
    Public Shared ReadOnly Property XMLPath As String
        Get
            Return String.Format("\\{0}\{1}\Versions", FlowServer.FlowServer, ShareDistribution)
        End Get
    End Property

    ' dgp rev 12/15/2010 
    Private Shared mShareDistribution As String = "Distribution"
    Public Shared ReadOnly Property ShareDistribution As String
        Get
            Return mShareDistribution
        End Get
    End Property

    Private Shared mCurName = "Configuration"

    Private Shared mScanList As Hashtable

    ' dgp rev 11/16/2010 
    Public Shared Function ModList(ByVal key As String, ByVal value As String) As Boolean

        ModList = False

        Dim Xele As XElement = XElement.Load(System.IO.Path.Combine(XMLPath, mCurName + ".xml"))

        For Each SubAtt In ConfigElements.Descendants.Attributes(key)
            ModList = True
            SubAtt.SetValue(value)
        Next

        Xele.Save(System.IO.Path.Combine(XMLPath, mCurName + ".xml"))

    End Function

    ' dgp rev 11/16/2010 
    Public Shared Function TestVersion() As String

        TestVersion = ""

        For Each SubAtt In ConfigElements.Descendants.Attributes("TestVersion")
            TestVersion = SubAtt.Value
        Next

    End Function

    ' dgp rev 8/1/2011 Username
    Private Shared mUserName As String = System.Environment.GetEnvironmentVariable("username")
    Public Shared ReadOnly Property UserName As String
        Get
            Return mUserName
        End Get
    End Property

    Private Shared mServerDistRoot = Nothing
    Private Shared Sub GetServerDistRoot()

        Dim Server = String.Format("\\{0}", mServer)

        Dim path = System.IO.Path.Combine(Server, "Distribution")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "Versions")
            If (System.IO.Directory.Exists(path)) Then
                If System.IO.Directory.GetDirectories(path).Length > 0 Then
                    mServerDistRoot = path
                End If
            End If
        End If

    End Sub

    ' dgp rev 6/24/09 Server Distribution Root
    Private Shared ReadOnly Property ServerDistRoot() As String
        Get
            If mServerDistRoot Is Nothing Then GetServerDistRoot()
            Return mServerDistRoot
        End Get
    End Property


    ' dgp rev 6/24/09 Retrieve Server Distribution List
    Private Shared Function GetServerDistList() As ArrayList

        GetServerDistList = New ArrayList

        If ServerDistRoot IsNot Nothing Then
            Dim item
            For Each item In System.IO.Directory.GetDirectories(ServerDistRoot)
                GetServerDistList.Add(System.IO.Path.GetFileName(item))
            Next
        End If

    End Function

    ' dgp rev 6/24/09 Server Distribution List
    Private Shared mServerDistList = Nothing
    Public Shared ReadOnly Property ServerDistList() As ArrayList
        Get
            If mServerDistList Is Nothing Then mServerDistList = GetServerDistList()
            Return mServerDistList
        End Get
    End Property

    ' dgp rev 6/23/09 Retrive the complete server distribution list
    Private Shared Function Check_Dist_Server() As Boolean

        Dim Server = String.Format("\\{0}", mServer)

        Check_Dist_Server = False
        Dim path = System.IO.Path.Combine(Server, "Distribution")
        If (System.IO.Directory.Exists(path)) Then
            path = System.IO.Path.Combine(path, "Versions")
            If (System.IO.Directory.Exists(path)) Then
                path = System.IO.Path.Combine(path, "Current")
                Check_Dist_Server = (System.IO.Directory.Exists(path))
            End If
        End If

    End Function

    ' dgp rev 6/24/09 Download Any Server Distribution to the Local Distribution
    Public Shared Function DownloadAnyServerDist() As String

        ' dgp rev 6/22/09 check server, then create a distribution and populate
        If (Check_Dist_Server()) Then
            ' dgp rev 7/16/08 download dist, if successful, set current
            If (DownloadAnyDist()) Then
                Return mDownloadedDist
            End If
        End If
        Return Nothing

    End Function

    Private Shared mServerAnyDist = Nothing
    Private Shared ReadOnly Property ServerAnyDist() As String
        Get
            If mServerAnyDist Is Nothing Then mServerAnyDist = Check_Dist_Server()
            Return mServerAnyDist
        End Get
    End Property

    Private Shared mDownloadedDist = Nothing
    ' dgp rev 7/15/08 Download the latest distribution
    Private Shared Function DownloadAnyDist() As Boolean

        If ServerDistList.Count = 0 Then Return False
        Dim index = ServerDistList.Count - 1
        If ServerDistList.Contains("Current") Then index = ServerDistList.IndexOf("Current")
        Dim tmp
        tmp = System.IO.Path.Combine(FlowStructure.FlowRoot, "Distribution")
        If (Not Utility.Create_Tree(tmp)) Then Return Nothing
        Try
            tmp = System.IO.Path.Combine(tmp, ServerDistList.Item(index))
            If (System.IO.Directory.Exists(tmp)) Then Return True
            Utility.DirectoryCopy(System.IO.Path.Combine(ServerDistRoot, ServerDistList.Item(index)), tmp)
            mDownloadedDist = tmp
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    ' dgp rev 7/15/08 Download the select distribution from server
    Public Shared Function DownloadSelectDist(ByVal name) As Boolean

        DownloadSelectDist = False
        Dim tmp
        tmp = System.IO.Path.Combine(FlowStructure.Dist_Root, name)
        If (Not HelperClasses.Utility.Create_Tree(tmp)) Then Exit Function
        If Not System.IO.Directory.Exists(tmp) Then Exit Function

        DownloadSelectDist = False
        If ServerDistList.Count > 0 Then
            Dim index
            If ServerDistList.Contains(name) Then
                index = ServerDistList.IndexOf(name)
                Try
                    Utility.DirectoryCopy(System.IO.Path.Combine(ServerDistRoot, name), tmp)
                    DownloadSelectDist = True
                Catch ex As Exception
                End Try
            End If
        End If

    End Function

    ' dgp rev 8/1/2011 Current Version  
    Private Shared mCurrentVersion = Nothing
    Private Shared mCurrentFlag = False
    Public Shared ReadOnly Property CurrentVersion
        Get
            If Not mCurrentFlag Then
                mCurrentFlag = True
                If PersonalFlag Then
                    mCurrentVersion = PersonalVersion
                Else
                    If DefaultFlag Then
                        mCurrentVersion = DefaultVersion
                    End If
                End If
            End If
            Return mCurrentVersion
        End Get
    End Property

    ' dgp rev 8/2/2011 Personal Version
    Private Shared mPersonalFlag As Boolean = False
    Private Shared mPersonalVersion As String = Nothing
    Public Shared ReadOnly Property PersonalVersion As String
        Get
            If mPersonalVersion Is Nothing Then PersonalScan()
            Return mPersonalVersion
        End Get
    End Property

    ' dgp rev 8/2/2011
    Public Shared ReadOnly Property PersonalFlag As Boolean
        Get
            If mPersonalVersion Is Nothing Then PersonalScan()
            Return mPersonalFlag
        End Get
    End Property

    ' dgp rev 11/16/2010 
    Private Shared Sub PersonalScan()

        mPersonalVersion = ""
        mPersonalFlag = False
        For Each SubXele In ConfigElements.Descendants(mUserName)
            For Each SubAtt In SubXele.Descendants.Attributes("CurVersion")
                mPersonalVersion = SubAtt.Value
                mPersonalFlag = Not mPersonalVersion = ""
            Next
        Next

    End Sub

    Private Shared mConfigElements As XElement
    Public Shared ReadOnly Property ConfigElements As XElement
        Get
            If mConfigElements Is Nothing Then mConfigElements = XElement.Load(System.IO.Path.Combine(XMLPath, mCurName + ".xml"))
            Return mConfigElements
        End Get
    End Property

    Private Shared SubAtt As Xml.Linq.XAttribute
    Private Shared SubXele As Xml.Linq.XElement

    ' dgp rev 8/2/2011 Personal Version
    Private Shared mDefaultFlag As Boolean = False
    Private Shared mDefaultVersion As String = Nothing
    Public Shared ReadOnly Property DefaultVersion As String
        Get
            If mDefaultVersion Is Nothing Then DefaultScan()
            Return mDefaultVersion
        End Get
    End Property

    ' dgp rev 8/2/2011
    Public Shared ReadOnly Property DefaultFlag As Boolean
        Get
            If mDefaultVersion Is Nothing Then DefaultScan()
            Return mDefaultFlag
        End Get
    End Property

    ' dgp rev 11/16/2010 
    Private Shared Sub DefaultScan()

        mDefaultVersion = ""
        mDefaultFlag = False
        For Each SubXele In ConfigElements.Descendants("default")
            For Each SubAtt In SubXele.Descendants.Attributes("CurVersion")
                mDefaultVersion = SubAtt.Value
                mDefaultFlag = Not mDefaultVersion = ""
            Next
        Next

    End Sub

    ' dgp rev 11/16/2010 
    Public Shared Function CheckList(ByVal key As String) As Boolean

        CheckList = False

        Dim item
        For Each item In ConfigElements.Descendants.Attributes(key)
            CheckList = True
        Next

    End Function

    ' dgp rev 11/16/2010 
    Public Shared Function ReadList(ByVal key As String) As String

        Dim conf_file As String
        ReadList = ""
        If System.IO.Directory.Exists(XMLPath) Then
            conf_file = System.IO.Path.Combine(XMLPath, mCurName + ".xml")
            If System.IO.File.Exists(conf_file) Then
                Dim Xele As XElement = XElement.Load(conf_file)
                Dim item
                ReadList = ""
                For Each item In Xele.Descendants.Attributes(key)
                    ReadList = item.value
                Next
            End If
        End If

    End Function

    ' dgp rev 11/16/2010 
    Public Shared Sub SaveList(ByVal CurVer)

        mScanList = New Hashtable
        mScanList.Add("CurVersion", CurVer)
        mScanList.Add("TestVersion", "YY")

        Dim node As XElement
        Dim xml = New XElement("FlowControl")
        Dim item As DictionaryEntry
        For Each item In mScanList
            node = New XElement(XName.Get("Settings"),
               New XAttribute(XName.Get(item.Key), XName.Get(item.Value)),
               New XElement(XName.Get("Keyword"), XName.Get("Value")))
            xml.Add(node)
        Next

        xml.Save(System.IO.Path.Combine(XMLPath, mCurName + ".xml"))

    End Sub

End Class
