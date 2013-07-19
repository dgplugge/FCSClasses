' Name:     PV-Wave Flow Control Session Class 
' Author:   Donald G Plugge
' Date:     7/17/07
' Purpose:  Controls the environment associated with a given project/session pair
' Description: Work with other FCS classes relating to data or other non-Project/Session 
' related variables, rather than incorporating that information into this class.
' Primary Goals:
' 1) Read the configuration file
' 2) Validate the previous working set including data list
' 3) Choose a default working set if none explicitly defined
' 4) Define a Work Root
' 5) Use external FlowRoot Class
Imports System.Xml
Imports System.IO
Imports FCS_Classes
Imports HelperClasses

Public Class Session

    Friend mMirrorFlowRoot As FlowRoot

    Private Descr_File As String = "Description.txt"
    Private Test_Run As Run_Name

    ' dgp rev 7/18/07 a local instance of the current run
    Private mRun As FCS_Classes.FCSRun
    Public Property Run() As FCS_Classes.FCSRun
        Get
            Return mRun
        End Get
        Set(ByVal value As FCS_Classes.FCSRun)
            mRun = value
        End Set
    End Property

    ' needed for the case when no session is assigned 6/8/07
    Private m_Off_Flag As Boolean = False
    Public Property Off_Flag() As Boolean
        Get
            Return m_Off_Flag
        End Get
        Set(ByVal value As Boolean)
            m_Off_Flag = value
        End Set
    End Property

    ' dgp rev 6/6/07 Runs that have been checked and listed in Valid or Missing
    Private m_Run_Status As Dictionary(Of String, Run_Name.CompStat)
    Public ReadOnly Property Run_Status() As Dictionary(Of String, Run_Name.CompStat)
        Get
            If (m_Run_Status Is Nothing) Then m_Run_Status = New Dictionary(Of String, Run_Name.CompStat)
            Return m_Run_Status
        End Get
    End Property
    ' dgp rev 6/6/07 Setting the data root then scan the all runs and initialize
    ' for future run checks - valid or missing

    ' dgp rev 5/17/07 Name of Data Run
    Private m_Data_Run As String
    Public Property Data_Run() As String
        Get
            If (m_Data_Path = "") Then Return ""
            If (Not System.IO.Directory.Exists(m_Data_Path)) Then Return ""
            Return System.IO.Path.GetDirectoryName(m_Data_Path)
        End Get
        Set(ByVal value As String)
            m_Data_Run = value
            If (m_Data_Path = "") Then Exit Property
            If (System.IO.Directory.Exists(m_Data_Path)) Then
                m_Data_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(m_Data_Path), value)
            End If
        End Set
    End Property
    ' dgp rev 5/17/07 Full Path to Data Run
    ' dgp rev 2/20/08 Sometimes there may be no data -- handle it.
    Private m_Data_Path As String
    Public Property Data_Path() As String
        Get
            If (m_Data_Path Is Nothing) Then
                If (Not FlowStructure.Data_Root Is Nothing) Then
                    Dim fld
                    If (System.IO.Directory.GetDirectories(FlowStructure.Data_Root).Length > 1) Then
                        For Each fld In System.IO.Directory.GetDirectories(FlowStructure.Data_Root)
                            m_Data_Path = fld
                            Exit For
                        Next
                    End If
                End If
            End If
            Return m_Data_Path
        End Get
        Set(ByVal value As String)
            m_Data_Path = value
            If (Not m_Data_Path Is Nothing) Then
                If (System.IO.Directory.Exists(m_Data_Path)) Then
                    m_Data_Run = System.IO.Path.GetDirectoryName(m_Data_Path)
                End If
            End If
        End Set
    End Property

    ' dgp rev 5/17/07 Data Root exists?
    Private m_Data_Root_Flag As Boolean = False
    Public ReadOnly Property Data_Root_Flag() As Boolean
        Get
            Return System.IO.Directory.Exists(FlowStructure.Data_Root)
        End Get
    End Property

    ' dgp rev 5/17/07 Data Root exists?
    Private m_Data_Path_Flag As Boolean = False
    Public ReadOnly Property Data_Path_Flag() As Boolean
        Get
            If (Data_Path Is Nothing) Then Return False
            Return System.IO.Directory.Exists(Data_Path)
        End Get
    End Property

    Private m_Data_Flag As Boolean = False

    ' dgp rev 5/23/07 set project and session to valid subdirectories
    Private Sub Default_Session()

        Dim proj, sess
        For Each proj In System.IO.Directory.GetDirectories(FlowStructure.Work_Root)
            For Each sess In System.IO.Directory.GetDirectories(System.IO.Path.GetDirectoryName(proj))
                SelectWork = System.IO.Path.GetDirectoryName(sess)
                Exit Sub
            Next
        Next

    End Sub

    ' dgp rev 5/23/07
    Private m_Valid As Boolean
    Public ReadOnly Property Valid() As Boolean
        Get
            Return m_Valid
        End Get
    End Property

    Private m_Project_Only As Boolean
    Public Property Project_Only() As Boolean
        Get
            Return m_Project_Only
        End Get
        Set(ByVal value As Boolean)
            m_Project_Only = value
        End Set
    End Property

    ' dgp rev 5/25/07 strip a path into project and session
    Private Sub Set_Proj_Sess(ByVal path As String)

        Dim arr() As String = path.ToLower.Split("\")
        Dim idx As Integer = Array.IndexOf(arr, "work")
        If (idx + 2 = arr.Length - 1) Then
            ' project and session
            m_Project = arr(idx + 1)
            m_Session = arr(idx + 2)
            'FlowStructure.Work_Root = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(path))
        Else
            m_Project = ""
            m_Session = ""
        End If

    End Sub

    ' dgp rev 5/25/07 Session and/or project have been found
    Private m_Found As Boolean
    Public ReadOnly Property Found() As Boolean
        Get
            Return m_Found
        End Get
    End Property

    ' dgp rev 6/7/07 check the status of the runs in the fcs file list
    Public Function Check_Runs() As Boolean

        Dim run As String = ""
        Dim local_run
        Dim results As Run_Name.CompStat = Run_Name.CompStat.Nomatch
        Dim mismatch_flag As Boolean = False
        Dim mismatch_run As String = ""

        Check_Runs = True

        ' list may contain more than one run
        For Each run In FCS_List.Unique_Runs
            ' for each run, see if the status already exists
            If (Run_Status.ContainsKey(run)) Then
                If (Run_Status.Item(run) <> Run_Name.CompStat.Match) Then Check_Runs = False
            Else
                Test_Run = New Run_Name(run)
                ' determine status by comparing run to each local run
                ' sort it into the appropriate list
                If (FlowStructure.RunsExist) Then
                    For Each local_run In FlowStructure.RunArray
                        results = Test_Run.Compare(local_run)
                        If (results = Run_Name.CompStat.Match) Then Exit For
                        If (results = Run_Name.CompStat.Mismatch) Then
                            mismatch_flag = True
                            mismatch_run = run
                        End If
                    Next
                End If
                ' if no match then check mismatch, otherwise take the results
                If (results = Run_Name.CompStat.Match) Then
                    Run_Status.Add(run, results)
                Else
                    Check_Runs = False
                    If (mismatch_flag) Then
                        Run_Status.Add(mismatch_run, Run_Name.CompStat.Mismatch)
                    Else
                        Run_Status.Add(run, results)
                    End If
                End If
            End If
        Next

    End Function

    ' dgp rev 9/27/2010


    ' dgp rev 5/23/07
    Private mSelectWork = Nothing
    Public Property SelectWork() As String
        Get
            If mSelectWork Is Nothing Then
                mSelectWork = FlowStructure.AnyWork
                Set_Proj_Sess(mSelectWork)
                ' read and validate the list
                m_FCS_List = New FCS_List(mSelectWork)
                ' validate each run
                m_Valid = m_FCS_List.AnyValid
            End If
            Return mSelectWork
        End Get
        Set(ByVal value As String)
            ' don't mess if nothing changes
            If (mSelectWork = value) Then Exit Property
            m_Found = System.IO.Directory.Exists(value)
            If (m_Found) Then
                mSelectWork = value
                Set_Proj_Sess(value)
                ' read and validate the list
                m_FCS_List = New FCS_List(value)
                ' validate each run
                m_Valid = m_FCS_List.AnyValid
                '                m_Valid = Validate()
            End If
        End Set
    End Property

    ' dgp rev 5/23/07
    Private m_Project As String
    Public ReadOnly Property Project() As String
        Get
            Return m_Project
        End Get
    End Property

    ' dgp rev 5/23/07
    Private ReadOnly Property Project_Path() As String
        Get
            If (Not Valid) Then Return ""
            Return System.IO.Path.GetDirectoryName(SelectWork)
        End Get
    End Property

    ' dgp rev 5/23/07
    Public Property Project_Desc() As String
        Get
            If (System.IO.File.Exists(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(SelectWork), Descr_File))) Then
                Dim sr As New System.IO.StreamReader(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(SelectWork), Descr_File))
                Dim line As String = sr.ReadToEnd
                sr.Close()
                Return line
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If (System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(SelectWork))) Then
                Dim sr As New System.IO.StreamWriter(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(SelectWork), Descr_File))
                sr.Write(value)
                sr.Close()
            End If
        End Set
    End Property

    ' dgp rev 5/23/07
    Public Property Session_Desc() As String
        Get
            If (System.IO.File.Exists(System.IO.Path.Combine(SelectWork, Descr_File))) Then
                Dim sr As New System.IO.StreamReader(System.IO.Path.Combine(SelectWork, Descr_File))
                Dim line As String = sr.ReadToEnd
                sr.Close()
                Return line
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If (System.IO.Directory.Exists(SelectWork)) Then
                Dim sr As New System.IO.StreamWriter(System.IO.Path.Combine(SelectWork, Descr_File))
                sr.Write(value)
                sr.Close()
            End If
        End Set
    End Property

    ' dgp rev 5/23/07
    Private m_Session As String
    Public ReadOnly Property Session() As String
        Get
            Return m_Session
        End Get
    End Property


    ' dgp rev 5/23/07
    Private m_Work_Root_Flag As Boolean = False
    Public ReadOnly Property Work_Root_Flag() As Boolean
        Get
            If (FlowStructure.Work_Root Is Nothing) Then Return False
            Return System.IO.Directory.Exists(FlowStructure.Work_Root)
        End Get
    End Property


    ' dgp rev 5/23/07
    Private m_Work_Flag As Boolean = False
    Public ReadOnly Property Work_Flag() As Boolean
        Get
            If (SelectWork Is Nothing) Then Return False
            Return System.IO.Directory.Exists(SelectWork)
        End Get
    End Property

    ' dgp rev 5/9/07 Scan Data Collection
    ' dgp rev 11/28/07 get the DataRoot from one place

    Public Function Scan_Runs() As Collection

        Dim tmp As New Collection
        Dim fld

        For Each fld In System.IO.Directory.GetDirectories(FlowStructure.Data_Root)
            tmp.Add(System.IO.Path.GetDirectoryName(fld), System.IO.Path.GetFileNameWithoutExtension(fld))
        Next
        Scan_Runs = tmp

    End Function

    ' dgp rev 5/18/07 Create Project and Session Readme, if missing
    Public Sub Create_Readme()

        If (System.IO.Directory.Exists(Project_Path)) Then
            Dim desc As String
            desc = Now.Date.ToString
            desc = desc + Data_Run
            desc = desc + System.IO.File.GetCreationTime(Project_Path).ToString
            Project_Desc = desc
            If (System.IO.Directory.Exists(SelectWork)) Then
                desc = Now.Date.ToString
                desc = desc + Data_Run
                desc = desc + System.IO.File.GetCreationTime((SelectWork)).ToString
                Session_Desc = desc
            End If
        End If

    End Sub

    ' dgp rev 5/18/07 Create Project and Session Readme, if missing
    Public Sub Check_Readme()

        Dim desc As String
        ' check for project path
        If (System.IO.Directory.Exists(Project_Path)) Then
            ' check project description
            If (Project_Desc Is Nothing) Then
                ' provide missing description
                desc = Now.Date.ToString
                desc = desc + vbCrLf + Data_Run
                desc = desc + vbCrLf + System.IO.File.GetCreationTime((Project_Path)).ToString
                Project_Desc = desc
            End If
            ' check for session path
            If (System.IO.Directory.Exists(SelectWork)) Then
                ' check session description
                If (Session_Desc Is Nothing) Then
                    ' provide missing description
                    desc = Now.Date.ToString
                    desc = desc + vbCrLf + Data_Run
                    desc = desc + vbCrLf + System.IO.File.GetCreationTime((SelectWork)).ToString
                    Session_Desc = desc
                End If
            End If
        End If
    End Sub

    Private m_Files As ArrayList
    Public ReadOnly Property Files() As ArrayList
        Get
            Return m_Files
        End Get
    End Property
    Private m_Missing_Files As ArrayList
    Private m_Runs As ArrayList
    Public ReadOnly Property Runs() As ArrayList
        Get
            Return m_Runs
        End Get
    End Property

    Private m_Run_Struct As Collection
    Public Property Run_Struct() As Collection
        Get
            Return m_Run_Struct
        End Get
        Set(ByVal value As Collection)
            m_Run_Struct = value
        End Set
    End Property

    Private m_Runs_Path As String

    Private m_Runs_Info As Run_Info

    '    Private Runs_Info As Dictionary(Of String, s_Run_Info)

    Private m_FCS_List As FCS_List

    Public Property FCS_List() As FCS_List
        Get
            Return m_FCS_List
        End Get
        Set(ByVal value As FCS_List)
            m_FCS_List = value
        End Set
    End Property

    Private m_Cluster_Flag As Boolean = True
    Public Property Cluster_Flag() As Boolean
        Get
            Return m_Cluster_Flag
        End Get
        Set(ByVal value As Boolean)
            m_Cluster_Flag = value
        End Set
    End Property
    Private m_Gate_Flag As Boolean = True
    Public Property Gate_Flag() As Boolean
        Get
            Return m_Gate_Flag
        End Get
        Set(ByVal value As Boolean)
            m_Gate_Flag = value
        End Set
    End Property
    Private m_Display_Flag As Boolean = True
    Public Property Display_Flag() As Boolean
        Get
            Return m_Display_Flag
        End Get
        Set(ByVal value As Boolean)
            m_Display_Flag = value
        End Set
    End Property

    ' dgp rev 5/25/07 Move the gates, clusters and other files
    Private Sub Move_Files(ByVal new_path As String)

        Dim file

        For Each file In System.IO.Directory.GetFiles(Me.SelectWork)
            If (System.IO.Path.GetFileName(file).ToLower.Contains("_mcs_file")) Then
                If (Cluster_Flag) Then System.IO.File.Copy(file, System.IO.Path.Combine(new_path, System.IO.Path.GetFileName(file)))
            ElseIf (System.IO.Path.GetFileName(file).ToLower.Contains("_mgs_struct")) Then
                If (Gate_Flag) Then System.IO.File.Copy(file, System.IO.Path.Combine(new_path, System.IO.Path.GetFileName(file)))
            ElseIf (System.IO.Path.GetFileName(file).ToLower.Contains("dmf_struct")) Then
                If (Display_Flag) Then System.IO.File.Copy(file, System.IO.Path.Combine(new_path, System.IO.Path.GetFileName(file)))
            Else
                System.IO.File.Copy(file, System.IO.Path.Combine(new_path, System.IO.Path.GetFileName(file)))
            End If
        Next

    End Sub

    ' dgp rev 5/23/07 New Instance
    Public Sub New()



    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
