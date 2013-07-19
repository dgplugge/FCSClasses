' Name:     FCS / Table Merger Routines
' Author:   Donald G Plugge
' Date:     3/22/06 
' Purpose:  Class to handle the merging of a tabular data set with each 
'           FCS file within a run
Imports System.io

Module MergeMod
    ' Rename flag

    Private m_DataSet As FCSRun
    Public Property objDataSet() As FCSRun
        Get
            Return m_DataSet
        End Get
        Set(ByVal value As FCSRun)
            m_DataSet = value
        End Set
    End Property
    Private m_table As FCSTable
    Public Property objTable() As FCSTable
        Get
            Return m_table
        End Get
        Set(ByVal value As FCSTable)
            m_table = value
        End Set
    End Property

    Private m_rename_flag As Boolean = True
    Public Property Rename_Flag() As Boolean
        Get
            Return m_rename_flag
        End Get
        Set(ByVal Value As Boolean)
            m_rename_flag = Value
        End Set
    End Property
    ' New FCS File Name
    Private m_new_name As String = objTable.env_date
    Public Property New_Name() As String
        Get
            Return m_new_name
        End Get
        Set(ByVal Value As String)
            m_new_name = Value
        End Set
    End Property
    ' The target path of the merged files
    Private m_target_path As String
    Public Property Target_Path() As String
        Get
            Return m_target_path
        End Get
        Set(ByVal Value As String)
            m_target_path = Value
        End Set
    End Property
    ' table extracted from FCS files 
    Private m_extract_table As ArrayList
    Public Property FCS_Table() As ArrayList
        Get
            Return m_extract_table
        End Get
        Set(ByVal Value As ArrayList)
            m_extract_table = Value
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

    ' dgp rev 3/23/06 Extract keys from the file object
    Private Function Keys_Extract(ByVal objFile As FCS_File) As ArrayList

        Dim item As String
        Dim Table_Info As New ArrayList

        For Each item In objFile.Header.Keys
            If (item(0) = "#") Then
                Table_Info.Add(item)
            End If
            If (item = "$SMNO") Then
                Table_Info.Add(item)
            End If
        Next

        Return Table_Info

    End Function

    ' dgp rev 3/23/06 Extract the key values from the current file object
    Private Function File_Extract(ByVal objFile As FCS_File) As ArrayList

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
    ' create a new fcs file with newly merged header
    Public Function Create_File(ByVal objFCS As FCS_File) As Boolean

        Dim new_path As String = MergeMod.Target_Path.ToString
        Dim new_name As String
        If (MergeMod.Rename_Flag) Then
            new_name = objTable.env_date + objFCS.Header("$SMNO") + ".FCS"
            objFCS.Header.Item("$FIL") = new_name
        Else
            new_name = Path.GetFileNameWithoutExtension(objFCS.Header("$FIL")) + ".FCS"
        End If
        Dim new_spec As String = Path.Combine(MergeMod.Target_Path.ToString, new_name)
        objFCS.Save_File(new_spec)

    End Function

    ' dgp rev 3/22/06 Merge the Table information into the FCS file headers
    Public Function Table_Merge() As Boolean

        Dim fcs_obj As FCS_File
        Dim keys As ArrayList
        Dim vals As ArrayList
        Dim idx As Int16

        keys = objTable.Get_Keys

        For idx = 0 To objDataSet.FCS_cnt - 1
            fcs_obj = objDataSet.FCS_Files(idx)
            vals = objTable.Get_Values(idx)
            If (File_Merge(fcs_obj, keys, vals)) Then
                If (Create_File(fcs_obj)) Then
                    Console.WriteLine("File created ")
                End If
            End If
        Next

    End Function

    ' dgp rev 3/22/06 Extract the Table information from the FCS file headers
    Public Function FCS_to_Matrix() As Boolean

        Dim Row As ArrayList
        Dim idx As Int16
        Dim holder As New ArrayList

        Row = Keys_Extract(objDataSet.FCS_Files(0))
        If (Row.Count > 0) Then holder.Add(Row)
        For idx = 0 To objDataSet.FCS_cnt - 1
            Row = File_Extract(objDataSet.FCS_Files(idx))
            If (Row.Count > 0) Then holder.Add(Row)
        Next
        m_extract_table = holder
        If (holder.Count > 0) Then Return True
        Return False

    End Function


End Module
