' Name:     FCS File Class
' Author:   Donald G Plugge
' Date:     07/17/07
' Purpose:  
Imports System.IO

Public Class FCS_List

    ' dgp rev 6/1/07 file list
    Private m_List_Spec As String
    ' dgp rev 9/29/2010 Make this property the only hard coding of FCS_Files.lis
    Public ReadOnly Property List_Spec() As String
        Get
            Return System.IO.Path.Combine(Work_Path, "fcs_files.lis")
        End Get
    End Property

    ' dgp rev 6/1/07 lines of list file
    Private m_List_Lines As Collection = Nothing
    Public ReadOnly Property List_Lines() As Collection
        Get
            If m_List_Lines Is Nothing Then
                m_List_Lines = New Collection
                If (System.IO.File.Exists(List_Spec)) Then
                    Dim sr = New StreamReader(List_Spec)
                    While (Not sr.EndOfStream)
                        m_List_Lines.Add(sr.ReadLine)
                    End While
                    sr.Close()
                End If
            End If
            Return m_List_Lines
        End Get
    End Property

    ' dgp rev 6/1/07 runs in file list
    Private m_Runs As Dictionary(Of String, Run_Info)
    Public Property Runs() As Dictionary(Of String, Run_Info)
        Get
            Return m_Runs
        End Get
        Set(ByVal value As Dictionary(Of String, Run_Info))
            m_Runs = value
        End Set
    End Property

    ' dgp rev 6/1/07 are any of the files invalid
    Public ReadOnly Property Valid() As Boolean
        Get
            Return (List_Lines.Count > 0)
        End Get
    End Property

    ' dgp rev 6/1/07 does the file list need to be rewritten
    Private m_Rewrite_Flag As Boolean = False
    Public Property Rewrite_Flag() As Boolean
        Get
            Return m_Rewrite_Flag
        End Get
        Set(ByVal value As Boolean)
            m_Rewrite_Flag = value
        End Set
    End Property

    ' dgp rev 6/1/07 missing files
    Private m_Missing As Collection
    Public Property Missing() As Collection
        Get
            Return m_Missing
        End Get
        Set(ByVal value As Collection)
            m_Missing = value
        End Set
    End Property

    ' dgp rev 6/1/07 work path 
    Private m_Work_Path As String
    Public ReadOnly Property Work_Path() As String
        Get
            Return m_Work_Path
        End Get
    End Property

    ' dgp rev 6/1/07 
    Public ReadOnly Property Unique_Runs() As ArrayList
        Get
            Return mListedRuns
        End Get
    End Property

    ' dgp rev 9/29/2010 "listed" exactly from the FCS file list
    Private mListedRuns As ArrayList
    Public ReadOnly Property ListedRuns() As ArrayList
        Get
            Return mListedRuns
        End Get
    End Property

    ' dgp rev 9/29/2010 "listed" exactly from the FCS file list
    Private mListedFileNames = Nothing
    Public ReadOnly Property ListedFileNames() As ArrayList
        Get
            Return mListedFileNames
        End Get
    End Property

    ' dgp rev 9/29/2010 Data root return the FCS List root unless it doesn't exist, then the global data root
    Public ReadOnly Property Data_Root() As String
        Get
            If m_Data_Root Is Nothing Then
                If AnyValid Then
                    Return mTrueDataRoot
                Else
                    Return FlowStructure.Data_Root
                End If
            Else
                Return m_Data_Root
            End If
        End Get
    End Property

    Public Sub Rename_Run(ByVal fromRun As String, ByVal toRun As String)

        Dim idx
        For idx = 0 To ListedRuns.Count - 1
            If (ListedRuns.Item(idx) = fromRun) Then ListedRuns.Item(idx) = toRun
        Next

        If (Unique_Runs.Contains(fromRun)) Then
            Unique_Runs.Remove(fromRun)
            Unique_Runs.Add(toRun)
        End If

    End Sub

    ' dgp rev 6/1/07 recreate the file list with the new data root
    Public Function RepairOneRun(ByVal run) As Boolean

        RepairOneRun = True
        Dim sw As New StreamWriter(List_Spec)
        Dim file
        For Each file In System.IO.Directory.GetFiles(System.IO.Path.Combine(Data_Root, run))
            sw.WriteLine(file)
        Next
        sw.Close()

    End Function

    ' dgp rev 6/1/07 recreate the file list with the new data root
    Public Function RepairByRun() As Boolean

        RepairByRun = True
        Dim sw As New StreamWriter(List_Spec)
        Dim run
        Dim file
        Dim file_test As String
        For Each run In Unique_Runs
            If TrueRuns.Contains(run) Then
                For Each file In TrueFiles
                    file_test = System.IO.Path.Combine(System.IO.Path.Combine(Data_Root, run), file)
                    If (System.IO.File.Exists(file_test)) Then sw.WriteLine(file_test)
                Next
            Else
                If RunFound(run) Then
                    For Each file In System.IO.Directory.GetFiles(System.IO.Path.Combine(Data_Root, run))
                        sw.WriteLine(file)
                    Next
                End If
            End If
        Next
        sw.Close()

    End Function

    ' dgp rev 6/1/07 recreate the file list with the new data root
    Public Function Recreate_List() As Boolean

        Dim sw As New StreamWriter(List_Spec)
        Dim idx As Integer
        Dim file_test As String
        For idx = 0 To ListedFileNames.Count - 1
            file_test = System.IO.Path.Combine(System.IO.Path.Combine(Data_Root, ListedRuns(idx)), ListedFileNames(idx))
            If (System.IO.File.Exists(file_test)) Then sw.WriteLine(file_test)
        Next
        sw.Close()

    End Function

    Private m_Data_Root = Nothing
    ' dgp rev 6/1/07 can we correct the error by simply changing the data root?
    ' dgp rev 11/28/07 get the DataRoot from one place
    Public Function Change_Data_Root(ByVal new_root As String) As Boolean

        Change_Data_Root = True ' assume success

        Dim idx As Integer
        Dim file_test As String
        For idx = 0 To ListedFileNames.Count - 1
            file_test = System.IO.Path.Combine(System.IO.Path.Combine(new_root, ListedRuns(idx)), ListedFileNames(idx))
            If (Not System.IO.File.Exists(file_test)) Then Change_Data_Root = False
        Next
        If (Change_Data_Root) Then m_Data_Root = new_root

    End Function

    ' dgp rev 9/29/2010 All files in list exist
    Public ReadOnly Property AllValid As Boolean
        Get
            Return TrueFiles.Count = ListedFileNames.Count
        End Get
    End Property

    ' dgp rev 9/29/2010 All files in list exist
    Public ReadOnly Property AnyValid As Boolean
        Get
            Return TrueFiles.Count > 0
        End Get
    End Property

    Private mTrueFiles = Nothing
    Public ReadOnly Property TrueFiles As ArrayList
        Get
            Return mTrueFiles
        End Get
    End Property

    ' dgp rev 9/29/2010 Missing Runs
    Private mMissingRuns = Nothing
    Public ReadOnly Property MissingRuns As ArrayList
        Get
            Return mMissingRuns
        End Get
    End Property

    Private mTrueDataRoot As String
    Private mTrueRuns = Nothing
    ' dgp rev 9/29/2010 True Runs meaning they exist exactly as specified in list
    Public ReadOnly Property TrueRuns As ArrayList
        Get
            Return mTrueRuns
        End Get
    End Property

    Private mMissingFiles As ArrayList

    ' dgp rev 9/30/2010 Missing run found, most likely the data root is invalid
    Private mRunFound As ArrayList
    Public ReadOnly Property RunFound(ByVal run_name) As Boolean
        Get
            Return mRunFound.Contains(run_name)
        End Get
    End Property


    ' dgp rev 6/1/07 Extract the data root from the first entry
    Public Sub ProcessList()

        mMissingRuns = New ArrayList
        mMissingFiles = New ArrayList
        mTrueFiles = New ArrayList
        mTrueRuns = New ArrayList
        mListedFileNames = New ArrayList
        mListedRuns = New ArrayList
        mRunFound = New ArrayList
        ' cycle thru the lines of the fcs list and check each file
        If (List_Lines.Count > 0) Then
            Dim item
            Dim file_name As String
            Dim run_name As String
            For Each item In List_Lines
                file_name = System.IO.Path.GetFileName(item)
                run_name = System.IO.Path.GetFileNameWithoutExtension(System.IO.Path.GetDirectoryName(item))
                If (System.IO.File.Exists(item)) Then
                    mTrueDataRoot = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(item))
                    mTrueFiles.Add(file_name)
                    If (Not mTrueRuns.Contains(run_name)) Then mTrueRuns.Add(run_name)
                Else
                    mMissingFiles.Add(file_name)
                    If (Not mMissingRuns.Contains(run_name)) Then mMissingRuns.Add(run_name)
                End If
                mListedFileNames.Add(file_name)
                If (Not mListedRuns.Contains(run_name)) Then
                    mListedRuns.Add(run_name)
                    If System.IO.Directory.Exists(System.IO.Path.Combine(Data_Root, run_name)) Then
                        mRunFound.Add(run_name)
                    End If
                End If
            Next
        End If

    End Sub

    Public ReadOnly Property MissingFiles As ArrayList
        Get
            Return mMissingFiles
        End Get
    End Property

    ' dgp rev 9/29/2010 how much smarts in the FCS List like finding missing files and run?
    Private Sub Init(ByVal work_path As String)

        m_Work_Path = work_path
        ProcessList()
        If MissingFiles.Count > 0 Then CorrectList()

    End Sub

    ' dgp rev 6/1/07 initialize the object by reading the list
    Public Sub New(ByVal work_path As String, ByVal data_root As String)

        Init(work_path)
        Change_Data_Root(data_root)

    End Sub

    ' dgp rev 6/1/07 initialize the object by reading the list
    Public Sub New(ByVal work_path As String)

        Init(work_path)

    End Sub


    Private mSelfCorrecting As Boolean = False
    Private Function CorrectList() As Boolean

        Dim item
        For Each item In MissingFiles

        Next

    End Function

    ' dgp rev 6/1/07 initialize the object by reading the list
    Public Sub New(ByVal work_path As String, ByVal SelfCorrecting As Boolean)

        mSelfCorrecting = SelfCorrecting
        Init(work_path)

    End Sub

End Class
