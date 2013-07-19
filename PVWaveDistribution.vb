
Imports Microsoft.Win32

' Name:     PV-Wave Distribution Handler
' Author:   Donald G Plugge
' Date:     2/6/2012
' Purpose:  Class to handle the PV-Wave distribution configuration

Public Class PVWaveDistribution

    Private Shared mDistributionServer = "NT-EIB-10-6B16"
    Private Shared mDistributionRemotePath = "Distribution\Versions"
    Private Shared mDistributionRemoteName = Nothing
    Private Shared mDistributionLocalName = Nothing
    Private Shared mDistributionLocalXML = Nothing

    Public Shared ReadOnly Property DistributionLocalName As String
        Get
            If mDistributionLocalName Is Nothing Then EstablishLocalPath()
            Return mDistributionLocalName
        End Get
    End Property

    Public Shared ReadOnly Property DistributionLocalXML As String
        Get
            Return mDistributionLocalXML
        End Get
    End Property

    Private Shared Sub EstablishLocalPath()

        Try
            mDistributionLocalName = ""
            If FlowStructure.ReadPersistantDist Then
                mDistributionLocalXML = FlowStructure.PersistantDist
                mDistributionLocalName = System.IO.Path.GetFileName(mDistributionLocalXML)
                If System.IO.Directory.Exists(mDistributionLocalXML) Then
                    FlowStructure.EstablishDistribution(mDistributionLocalXML)
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    ' dgp rev 6/24/09 PVW Path
    Public Shared ReadOnly Property DistributionLocalSpec() As String
        Get
            If mDistributionLocalName Is Nothing Then EstablishLocalPath()
            Return System.IO.Path.Combine(FlowStructure.Dist_Root, DistributionLocalName)
        End Get
    End Property

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property DistributionRemoteSpec As String
        Get
            Return System.IO.Path.Combine(ServerRoot, DistributionRemoteName)
        End Get
    End Property

    ' dgp rev 5/23/2012 Distribution Remote Name based upon username
    Public Shared ReadOnly Property DistributionRemoteName As String
        Get
            If mDistributionRemoteName Is Nothing Then
                If XMLStore.Exists(Environment.UserName, "DistributionName") Then
                    mDistributionRemoteName = XMLStore.GetValue(Environment.UserName, "DistributionName")
                Else
                    If XMLStore.Exists("DistributionName") Then
                        mDistributionRemoteName = XMLStore.GetValue("DistributionName")
                    Else
                        mDistributionRemoteName = "Current"
                        If XMLStore.AddValue("DistributionName", mDistributionRemoteName) Then
                            mDistributionRemoteName = XMLStore.GetValue("DistributionName")
                        End If
                    End If
                End If
            End If
            Return mDistributionRemoteName
        End Get
    End Property
    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property DistributionRemotePath As String
        Get
            Return mDistributionRemotePath
        End Get
    End Property

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property ServerRoot As String
        Get
            Return String.Format("\\{0}\{1}", mDistributionServer, mDistributionRemotePath)
        End Get
    End Property

    Private Shared mXMLStore = Nothing
    Private Shared ReadOnly Property XMLStore As HelperClasses.XMLStore
        Get
            If mXMLStore Is Nothing Then
                mXMLStore = New HelperClasses.XMLStore(ServerRoot, "PVWaveTest")
            End If
            Return mXMLStore
        End Get
    End Property


    ' dgp rev 2/3/2012
    Private Shared mDyn = Nothing
    Private Shared ReadOnly Property Dyn As HelperClasses.Dynamic
        Get
            If mDyn Is Nothing Then
                mDyn = New HelperClasses.Dynamic(ServerRoot, "PVWave")
            End If
            Return mDyn
        End Get
    End Property

    ' dgp rev 2/3/2012
    Private Shared Function DistributionRemoteExists() As Boolean

        DistributionRemoteExists = False
        If FlowServer.Server_Up Then
            If System.IO.Directory.Exists(ServerRoot) Then
                Return System.IO.Directory.Exists(System.IO.Path.Combine(ServerRoot, DistributionRemoteName))
            End If
        End If

    End Function

    Private Shared mDistributionRemoteFlag As Boolean = False
    Private Shared mDistributionLocalFlag As Boolean = False

    ' dgp rev 2/6/2012 Compare Remote and Local names to assure a match
    ' assume both files exist
    Private Shared Function CompareNames() As Boolean

        Return (DistributionRemoteName.ToLower = DistributionLocalName.ToLower)

    End Function

    ' dgp rev 2/6/2012 Sync Local with Remote, transfer Remote to Local
    ' assume that the Remote exists
    Private Shared Function SyncDistribution() As Boolean

        Try
            mDistributionLocalName = DistributionRemoteName
            If Not DistributionServer.DownloadSelectDist(DistributionRemoteName) Then Return False
            FlowStructure.UpdatePersistantDist(DistributionLocalSpec)
            FlowStructure.EstablishDistribution(DistributionLocalSpec)
            Return (FlowStructure.PersistantDist = DistributionLocalSpec)

        Catch ex As Exception
            Return False
        End Try

    End Function

    ' dgp rev 2/3/2012
    Public Shared Function ValidDistributionFound() As Boolean

        ValidDistributionFound = True
        If DistributionRemoteExists() Then
            If DistributionLocalExists() Then
                If Not CompareNames() Then SyncDistribution()
            Else
                SyncDistribution()
            End If
        Else
            Return DistributionLocalExists()
        End If

    End Function

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property DistributionLocalExists As Boolean
        Get
            If DistributionLocalName = "" Then Return False
            Return System.IO.Directory.Exists(DistributionLocalXML)
        End Get
    End Property

End Class
