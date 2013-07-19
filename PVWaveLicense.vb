Imports Microsoft.Win32
Imports System.IO

' Name:     PV-Wave License Handler
' Author:   Donald G Plugger
' Date:     2/3/2012
' Purpose:  Class to handle the PV-Wave license configuration
Public Class PVWaveLicense

    Private Shared mLicenseServer = "NT-EIB-10-6B16"
    Private Shared mLicenseRemotePath = "Distribution\PVW License"
    Private Shared mLicenseRemoteName = Nothing
    Private Shared mLicenseLocalName = Nothing
    Private Shared mLicenseLocalReg = Nothing

    Private Shared mLicenseRegPath As String = "Environment"
    Private Shared mLicenseEnvVar As String = "LM_LICENSE_FILE"

    Public Shared ReadOnly Property LicenseLocalName As String
        Get
            If mLicenseLocalName Is Nothing Then EstablishLocalPath()
            Return mLicenseLocalName
        End Get
    End Property

    Private Shared mLicensePath As String = FlowStructure.Dist_Root
    Public Shared Property LicensePath As String
        Get
            Return mLicensePath
        End Get
        Set(ByVal value As String)
            mLicensePath = value
        End Set
    End Property

    Private Shared Sub EstablishLocalPath()

        mLicenseLocalName = "License.Lic"
        mLicenseLocalReg = System.IO.Path.Combine(LicensePath, mLicenseLocalName)
        Try
            If (ReadReg()) Then
                If System.IO.Path.GetDirectoryName(mLicenseLocalReg).ToLower = LicensePath.ToLower Then
                    mLicenseLocalName = System.IO.Path.GetFileName(mLicenseLocalReg)
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Shared Function ReadReg() As Boolean
        Dim RegKey As RegistryKey
        ReadReg = False
        mLicenseLocalReg = ""
        Try
            RegKey = Registry.CurrentUser.OpenSubKey(mLicenseRegPath, False)
            If (RegKey Is Nothing) Then
                RegKey.CreateSubKey(mLicenseRegPath)
            End If
            mLicenseLocalReg = RegKey.GetValue(mLicenseEnvVar)
            FlowStructure.Log_Info(String.Format("Current License {0}", mLicenseLocalReg))
            ReadReg = True

        Catch ex As Exception

        End Try

    End Function

    ' dgp rev 6/24/09 PVW Path
    Public Shared ReadOnly Property LicenseLocalSpec() As String
        Get
            If mLicenseLocalName Is Nothing Then EstablishLocalPath()
            Return System.IO.Path.Combine(LicensePath, LicenseLocalName)
        End Get
    End Property



    ' dgp rev 6/24/09 Check registry for PV-Wave path
    Private Shared Function DefineRegistryKey(ByVal local As String) As Boolean

        Dim RegKey As RegistryKey
        DefineRegistryKey = False

        Try
            RegKey = Registry.CurrentUser.OpenSubKey(mLicenseRegPath, True)
            If (RegKey Is Nothing) Then
                RegKey.CreateSubKey(mLicenseRegPath)
            End If
            RegKey.SetValue(mLicenseEnvVar, local)
            FlowStructure.Log_Info(String.Format("Defined License {0}", local))
            DefineRegistryKey = True
        Catch ex As Exception
            FlowStructure.Log_Info(String.Format("License Error {0}", ex.Message))
        End Try

    End Function

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property LicenseRemoteSpec As String
        Get
            Return System.IO.Path.Combine(ServerRoot, LicenseRemoteName)
        End Get
    End Property

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property LicenseRemoteName As String
        Get
            If mLicenseRemoteName Is Nothing Then
                If Dyn.Exists("LicenseName") Then
                    mLicenseRemoteName = Dyn.GetSetting("LicenseName")
                Else
                    mLicenseRemoteName = "License.Lic"
                    Dyn.PutSetting("LicenseName", mLicenseRemoteName)
                End If
            End If
            Return mLicenseRemoteName
        End Get
    End Property
    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property LicenseRemotePath As String
        Get
            Return mLicenseRemotePath
        End Get
    End Property

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property ServerRoot As String
        Get
            Return String.Format("\\{0}\{1}", mLicenseServer, mLicenseRemotePath)
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
    Private Shared Function LicenseRemoteExists() As Boolean

        LicenseRemoteExists = False
        If FlowServer.Server_Up Then
            If System.IO.Directory.Exists(ServerRoot) Then
                Return System.IO.File.Exists(System.IO.Path.Combine(ServerRoot, LicenseRemoteName))
            End If
        End If

    End Function

    Private Shared mLicenseRemoteFlag As Boolean = False
    Private Shared mLicenseLocalFlag As Boolean = False

    ' dgp rev 2/6/2012 Compare Remote and Local names to assure a match
    ' assume both files exist
    Private Shared Function CompareNames() As Boolean

        Return (LicenseRemoteName.ToLower = LicenseLocalName.ToLower)

    End Function

    ' dgp rev 2/6/2012 Sync Local with Remote, transfer Remote to Local
    ' assume that the Remote exists

    Private Shared Function RemoveExistingLicenseFiles() As Boolean

        Dim dirinfo As DirectoryInfo = New DirectoryInfo(System.IO.Path.GetDirectoryName(LicenseLocalSpec))
        Dim mAllFiles = dirinfo.GetFiles("*.lic", SearchOption.TopDirectoryOnly)
        Dim item As FileInfo
        Try
            For Each item In mAllFiles
                System.IO.File.Delete(item.FullName)
            Next
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Private Shared Function SyncLicense() As Boolean

        Try
            mLicenseLocalName = LicenseRemoteName
            RemoveExistingLicenseFiles()
            System.IO.File.Copy(LicenseRemoteSpec, LicenseLocalSpec, True)
            Return DefineRegistryKey(LicenseLocalSpec)

        Catch ex As Exception
            Return False
        End Try

    End Function

    ' dgp rev 2/3/2012
    Public Shared Function ValidLicenseFound() As Boolean

        ValidLicenseFound = True
        If LicenseRemoteExists() Then
            If LicenseLocalExists() Then
                If Not CompareNames() Then SyncLicense()
            Else
                SyncLicense()
            End If
        Else
            Return LicenseLocalExists()
        End If

    End Function

    ' dgp rev 2/3/2012
    Public Shared ReadOnly Property LicenseLocalExists As Boolean
        Get
            Return System.IO.File.Exists(LicenseLocalSpec)
        End Get
    End Property

End Class
