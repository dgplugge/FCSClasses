Imports System.IO
Imports HelperClasses

' dgp rev 6/5/07 Create a new working project and session
Public Class New_Work


    Private m_project As String
    Private m_session As String
    Private m_root As String
    Private m_new_session As String = ""
    Public ReadOnly Property New_Session() As String
        Get
            Return m_new_session
        End Get
    End Property

    Private m_create_flag As Boolean = False

    ' dgp rev 6/5/07 session, project, root defined
    Private m_ready_flag As Boolean = False
    Public ReadOnly Property Ready_Flag() As Boolean
        Get
            Return m_ready_flag
        End Get
    End Property

    ' dgp rev 6/5/07 path to root of runs
    Private m_Data_Root As String
    Public ReadOnly Property Data_Root() As String
        Get
            Return m_Data_Root
        End Get
    End Property

    ' dgp rev 6/5/07 list of all runs, initialize to empty list
    Private m_Run_List As New ArrayList
    Public ReadOnly Property Run_List() As ArrayList
        Get
            Return m_Run_List
        End Get
    End Property

    ' dgp rev 6/11/07
    Private m_gate_file As New ArrayList
    Public ReadOnly Property Gate_File() As ArrayList
        Get
            Return m_gate_file
        End Get
    End Property

    ' dgp rev 6/11/07
    Private m_gate_flag As Boolean
    Public Property Gate_Flag() As Boolean
        Get
            Return m_gate_flag
        End Get
        Set(ByVal value As Boolean)
            m_gate_flag = value
        End Set
    End Property

    ' dgp rev 6/11/07
    Private m_cluster_file As New ArrayList
    Public ReadOnly Property Cluster_File() As ArrayList
        Get
            Return m_cluster_file
        End Get
    End Property

    ' dgp rev 6/11/07
    Private m_cluster_flag As Boolean
    Public Property Cluster_Flag() As Boolean
        Get
            Return m_cluster_flag
        End Get
        Set(ByVal value As Boolean)
            m_cluster_flag = value
        End Set
    End Property

    ' dgp rev 6/11/07
    Private m_display_file As New ArrayList
    Public ReadOnly Property Display_File() As ArrayList
        Get
            Return m_display_file
        End Get
    End Property

    ' dgp rev 6/11/07
    Private m_display_flag As Boolean
    Public Property Display_Flag() As Boolean
        Get
            Return m_display_flag
        End Get
        Set(ByVal value As Boolean)
            m_display_flag = value
        End Set
    End Property

    Private m_sess_desc As String
    Private m_proj_desc As String
    ' dgp rev 6/5/07 new project
    Private m_new_flag As Boolean
    Public ReadOnly Property New_Flag() As Boolean
        Get
            Return m_new_flag
        End Get
    End Property
    Private m_base_work As String
    ' dgp rev 6/5/07 is the project valid
    Private m_valid As Boolean = False
    Public ReadOnly Property Valid() As Boolean
        Get
            Return m_valid
        End Get
    End Property

    ' dgp rev 4/12/07 Create a Unique Name
    Public Function Unique_Name() As String

        Return Format(Now(), "yyyyMMddhhmmss")

    End Function
    ' dgp rev 6/5/07 move proper files from previous session to new session
    Private Sub Move_Files()

        If (m_create_flag) Then

            Dim item
            Dim target As String
            If (m_cluster_flag And m_cluster_file.Count > 0) Then
                For Each item In m_cluster_file
                    target = System.IO.Path.Combine(m_new_session, System.IO.Path.GetFileName(item))
                    System.IO.file.copy(item, target)
                Next
            End If
            If (m_gate_flag And m_gate_file.Count > 0) Then
                For Each item In m_gate_file
                    target = System.IO.Path.Combine(m_new_session, System.IO.Path.GetFileName(item))
                    System.IO.file.copy(item, target)
                Next
            End If
            If (m_display_flag And m_display_file.Count > 0) Then
                For Each item In m_display_file
                    target = System.IO.Path.Combine(m_new_session, System.IO.Path.GetFileName(item))
                    System.IO.file.copy(item, target)
                Next
            End If

        End If

    End Sub

    ' dgp rev 6/5/07 return the full path of the new session.
    Private Sub Create_Session()

        If (Valid) Then
            If (System.IO.Directory.exists(m_root)) Then
                m_new_session = system.io.path.combine(system.io.path.combine(m_root, m_project), m_session)
                m_create_flag = Utility.Create_Tree(m_new_session)
            End If
        End If

    End Sub
    ' dgp rev 6/5/07 create of list of the data files
    Private Sub Create_Data_List()

        Dim objFCSList As FCS_List = New FCS_List(m_new_session)
        Dim sw As New StreamWriter(objFCSList.List_Spec)

        Dim item
        Dim file
        Dim run_path
        For Each item In Run_List
            run_path = System.IO.Path.Combine(Data_Root, item)
            For Each file In System.IO.Directory.GetFiles(run_path)
                sw.WriteLine(file.ToString)
            Next
        Next
        sw.Close()

    End Sub

    ' dgp rev 6/12/07 Create a new session description
    Public Sub New_Sess_Desc()

        Dim txt As String
        txt = "Session " + m_session + vbCrLf
        txt = txt + "Created on " + CStr(Now()) + vbCrLf
        Session_Desc = txt

    End Sub

    ' dgp rev 6/12/07 Create a new project description
    Public Sub New_Proj_Desc()

        Dim txt As String
        txt = "Project " + m_project + vbCrLf
        txt = txt + "Created on " + CStr(Now()) + vbCrLf
        Project_Desc = txt

    End Sub
    ' dgp rev 6/11/07 Create an initial project description
    Private Sub Create_Descripts()

        Dim file_spec As String
        Dim sw As System.IO.StreamWriter

        Dim path As String
        ' new project, new description, so create it
        path = System.IO.Path.Combine(m_root, m_project)
        file_spec = System.IO.Path.Combine(path, "Description.txt")
        sw = New System.IO.StreamWriter(file_spec)
        sw.Write(Me.Project_Desc)
        sw.Close()

        ' new project, new session, new description, create it
        file_spec = System.IO.Path.Combine(m_new_session, "Description.txt")
        sw = New System.IO.StreamWriter(file_spec)
        sw.Write(Session_Desc)
        sw.Close()

    End Sub

    ' dgp rev 6/5/07 Create the new session
    Public Sub Create_New_Session()

        Create_Session()
        If (m_create_flag) Then
            Move_Files()
            Create_Data_List()
            Create_Descripts()
        End If

    End Sub

    ' dgp rev 6/11/07 Check Path
    Private Function Check_Path(ByVal path As String) As Boolean

        If (Not System.IO.Directory.Exists(path)) Then Return False

        Dim arr() As String = path.Split("\")

        ' dgp rev 5/13/09 new path
        If (arr(arr.Length - 1).ToLower = "work") Then
            m_new_flag = True
            m_project = Unique_Name()
            m_session = m_project
            m_root = path
        Else
            If (arr.Length > 3) Then
                If (arr(arr.Length - 3).ToLower <> "work") Then Return False
                m_base_work = path
                m_project = arr(arr.Length - 2)
                m_session = Unique_Name()
                Dim idx
                m_root = arr(0) + "\"
                For idx = 1 To arr.Length - 3
                    m_root = System.IO.Path.Combine(m_root, arr(idx))
                Next
            End If
        End If

        Return (System.IO.Directory.Exists(m_root))

    End Function
    ' dgp rev 6/11/07 Check Data 
    Private Function Check_Data(ByVal path As String) As Boolean

        If (Not System.IO.Directory.Exists(path)) Then Return False
        Check_Data = False

        Dim arr() As String = path.Split("\")

        If (arr.Length > 2) Then
            If (arr(arr.Length - 2).ToLower <> "data") Then Return False
            Dim run_name As String = System.IO.Path.GetFileName(path)
            If (Run_List.Contains(run_name)) Then Return False
        End If

        Return (System.IO.Directory.Exists(m_root))

    End Function

    ' dgp rev 6/5/07 clear the data list
    Public Sub Clear_Datset()

        m_Run_List.Clear()
        ' work is no longer valid
        m_valid = False

    End Sub

    ' dgp rev 6/5/07 append data to the current data list
    Public Function Add_Dataset(ByVal path As String) As Boolean

        If (Check_Data(path)) Then
            If (m_Data_Root Is Nothing) Then m_Data_Root = system.io.Path.GetDirectoryName(path)
            Run_List.Add(System.IO.Path.GetFileName(path))
        End If
        m_valid = (Ready_Flag And Run_List.Count > 0)

    End Function

    'dgp rev 6/12/07
    Public Sub Sess_Desc_Append(ByVal text As String)

        Session_Desc = text + vbCrLf + "Appended on " + Now() + vbCrLf

    End Sub

    ' dgp rev 5/23/07
    Public Property Session_Desc() As String
        Get
            Return m_sess_desc
        End Get
        Set(ByVal value As String)
            m_sess_desc = value
        End Set
    End Property

    ' dgp rev 5/23/07
    Public Property Project_Desc() As String
        Get
            Return m_proj_desc
        End Get
        Set(ByVal value As String)
            m_proj_desc = value
        End Set
    End Property

    ' dgp rev 6/5/07 Scan work for cluster, gate and display files
    Private Sub Scan_Work()

        m_gate_flag = False
        m_cluster_flag = False
        m_display_flag = False

        m_cluster_file.Clear()
        m_gate_file.Clear()
        m_display_file.Clear()

        Dim file

        For Each file In System.IO.Directory.GetFiles(m_base_work)
            If (System.IO.Path.GetFileName(file).ToLower.Contains("_mcs_file")) Then
                m_cluster_file.Add(file)
            End If
            ' dgp rev 2/9/2011 add the car structure
            If (System.IO.Path.GetFileName(file).ToLower.Contains("car_struct")) Then
                m_cluster_flag = True
                m_cluster_file.Add(file)
            End If
            If (System.IO.Path.GetFileName(file).ToLower.Contains("_mgs_struct")) Then
                m_gate_flag = True
                m_gate_file.Add(file)
            End If
            If (System.IO.Path.GetFileName(file).ToLower.Contains("dmf_struct")) Then
                m_display_flag = True
                m_display_file.Add(file)
            End If
        Next

    End Sub

    ' dgp rev 6/5/07 Create a new instance from previous session
    Public Sub New(ByVal Path As String)

        If (Check_Path(Path)) Then
            m_ready_flag = True
            ' project, session and root successfully defined
            If (New_Flag) Then
            Else
                Scan_Work()
                m_sess_desc = system.io.path.combine(m_base_work, "description.txt")
                m_proj_desc = system.io.path.combine(system.io.path.combine(m_root, m_project), "description.txt")
            End If
        End If

    End Sub

End Class
