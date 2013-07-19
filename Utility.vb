Public Class Utility

    Shared fso As Scripting.FileSystemObject

    ' dgp rev 11/8/06 create a folder with error handling
    ' dgp rev 11/13/06 parent folder must exist
    Private Shared Function create_folder(ByVal strPath As String)

        create_folder = False ' assume failure
        If (Not fso.FolderExists(strPath)) Then
            On Error GoTo Create_Error
            If (fso.FolderExists(fso.GetParentFolderName(strPath))) Then
                fso.CreateFolder(strPath)
            End If
        End If
        create_folder = True
        Exit Function

Create_Error:
        MsgBox("Unable to create folder -- " + strPath, vbExclamation, "Folder Creation Error")

    End Function

    ' dgp rev 11/27/06 once it is determined that a path is to be created, back up
    ' to the first subdirectory that exists.
    Public Shared Function Create_Tree(ByVal path_str As String) As Boolean

        Dim test_path As String
        Dim sep As Char = System.IO.Path.DirectorySeparatorChar
        Dim split_arr() As String = path_str.Split(sep)

        If (fso.FolderExists(path_str)) Then Return True

        Dim idx
        test_path = split_arr(0) + sep
        If (Not fso.FolderExists(test_path)) Then Return False

        For idx = 1 To split_arr.Length - 1
            test_path = test_path + split_arr(idx) + sep
            If (Not fso.FolderExists(test_path)) Then create_folder(test_path)
        Next

        If (fso.FolderExists(path_str)) Then Return True

    End Function

End Class
