' Name:     FCS File Class
' Author:   Donald G Plugge
' Date:     3/17/06 
' Purpose:  Class to handle reading, writing and manipulating FCS files

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Public Class FCS_File

    ' If invalid, why?
    Private m_invalid As String
    Public Property InValid() As String
        Get
            Return m_invalid
        End Get
        Set(ByVal Value As String)
            m_invalid = Value
        End Set
    End Property
    ' has table information
    Private m_table_flag As Boolean
    Public ReadOnly Property Table_Flag() As Boolean
        Get
            If (m_header Is Nothing) Then Read_Header()
            Return m_table_flag
        End Get
    End Property
    ' FCS Version
    Private m_version As String
    Public Property Version() As String
        Get
            Return m_version
        End Get
        Set(ByVal Value As String)
            m_version = Value
        End Set
    End Property

    ' Is a valid FCS file
    Private m_valid As Boolean
    Public Property Valid() As Boolean
        Get
            Return m_valid
        End Get
        Set(ByVal Value As Boolean)
            m_valid = Value
        End Set
    End Property

    ' Is a valid FCS file
    Private mValidData = Nothing
    Public ReadOnly Property ValidData() As Boolean
        Get
            If m_datablock Is Nothing Then Read_Data()
            Return m_valid
        End Get
    End Property

    ' Path to the FCS file
    Private m_path As String
    Public Property Path() As String
        Get
            Return m_path
        End Get
        Set(ByVal Value As String)
            m_path = Value
        End Set
    End Property

    ' full file specification
    Private m_fullspec As String
    Public Property FullSpec() As String
        Get
            Return m_fullspec
        End Get
        Set(ByVal Value As String)
            m_fullspec = Value
        End Set
    End Property

    ' dgp rev 8/6/2010 TubeNumber 
    Public Function TubeNumber(ByRef Num As Int16) As Boolean

        Num = -1
        If (Header.ContainsKey("TUBE NAME")) Then
            Dim Val = Header("TUBE NAME")
            Dim arr = Val.Split("_")
            Dim isNum = Long.TryParse(arr(arr.length - 1), Num)
            Return isNum
        End If
        Return False

    End Function

    ' file name and extension
    Private m_filename As String
    Public Property FileName() As String
        Get
            Return m_filename
        End Get
        Set(ByVal Value As String)
            m_filename = Value
        End Set
    End Property

    Private m_keep_open As Boolean = False

    Private m_fixed_size As Int16 = 59
    ' 10 character fixed header containing FCS version number
    Private m_Ver(10) As Char
    ' Version number
    Private m_vernum As Int16

    ' Fixed Header Pointers
    Private m_pointers(6) As Int32
    Private m_RevisedPointers(6) As Int32

    ' dgp rev 4/21/09
    Private mProtocolCheck As Boolean = False
    Private mProtocolExists As Boolean = False
    Public ReadOnly Property ProtocolExists() As Boolean
        Get
            If (Not mProtocolCheck) Then mProtocolKeys = ExtractProtocolKeys()
            Return mProtocolExists
        End Get
    End Property

    Private mProtocolKeys As ArrayList
    Public ReadOnly Property ProtocolKeys() As ArrayList
        Get
            If (Not ProtocolExists) Then
                mProtocolKeys = New ArrayList
            End If
            Return mProtocolKeys
        End Get
    End Property

    Private mProtocolVals As ArrayList
    ' dgp rev 4/20/09 Return the Protocol Values
    Public Function ProtocolValues(ByVal keys As ArrayList) As ArrayList

        Dim item
        Dim tmp = New ArrayList
        For Each item In keys
            If (Header.ContainsKey(item)) Then
                tmp.Add(Header.Item(item))
            Else
                tmp.Add("N/A")
            End If
        Next
        Return tmp

    End Function



    ' dgp rev 4/20/09 Return the Protocol Values
    Public Function ProtocolValues() As ArrayList

        Dim item
        mProtocolVals = New ArrayList
        mProtocolVals.Add(Header.Item("$SMNO"))
        For Each item In Me.ProtocolOrder
            If (Header.ContainsKey(item)) Then
                mProtocolVals.Add(Header.Item(item))
            Else
                mProtocolVals.Add(" ")
            End If
        Next
        Return mProtocolVals

    End Function

    ' dgp rev 3/23/06 Extract keys from the file object
    Public Function ExtractProtocolKeys() As ArrayList

        Dim item As String
        Dim Table_Info As New ArrayList

        For Each item In Header.Keys
            If (item(0) = "#") Then
                Table_Info.Add(item)
            End If
            If (item = "$SMNO") Then
                Table_Info.Add(item)
            End If
        Next

        mProtocolCheck = True
        mProtocolExists = (Table_Info.Count > 0)
        Return Table_Info

    End Function



    Public Property RevisedPointers(ByVal idx As Int16) As Int32
        Get
            Return m_RevisedPointers(idx)
        End Get
        Set(ByVal Value As Int32)
            m_RevisedPointers(idx) = Value
        End Set
    End Property

    Public Property Pointers(ByVal idx As Int16) As Int32
        Get
            Return m_pointers(idx)
        End Get
        Set(ByVal Value As Int32)
            m_pointers(idx) = Value
        End Set
    End Property

    ' Fixed Header Block as character array
    Private m_fixed As Byte()
    Public Property Fixed() As Byte()
        Get
            Return m_fixed
        End Get
        Set(ByVal Value As Byte())
            m_fixed = Value
        End Set
    End Property

    ' Text Block as character array
    Private m_textblock As Char()
    Public Property Text_Block() As Char()
        Get
            Return m_textblock
        End Get
        Set(ByVal Value As Char())
            m_textblock = Value
        End Set
    End Property

    ' Separator character
    Private m_separator As Char
    Public Property Separator() As Char
        Get
            Return m_separator
        End Get
        Set(ByVal value As Char)
            m_separator = value
        End Set
    End Property

    ' Text Header as broken into hash
    Private m_header As Hashtable

    Private mProtOrder As ArrayList
    Public ReadOnly Property ProtocolOrder() As ArrayList
        Get
            Return mProtOrder
        End Get
    End Property

    ' Read in the text header
    Private Sub Read_Header()

        Dim parse_str As String
        Dim str_arr() As String
        Dim header_hash As New Hashtable
        Dim txt_len As Integer = Pointers(1) - Pointers(0) + 1
        Dim block_len As Integer = Pointers(1)

        Dim block(block_len) As Char

        Dim txt_Reader As New StreamReader(FullSpec)

        If (txt_Reader.ReadBlock(block, 0, block_len) = block_len) Then
            ReDim m_textblock(txt_len)
            Array.Copy(block, Pointers(0), m_textblock, 0, txt_len)
            parse_str = m_textblock
            m_separator = m_textblock(0)
            m_table_flag = Regex.IsMatch(m_textblock, m_separator + "#")
            mProtOrder = New ArrayList
            If (m_table_flag) Then
                Dim info = Regex.Matches(m_textblock, m_separator + "(#.+?)" + m_separator)
                Dim item
                For Each item In info
                    Try
                        mProtOrder.Add(item.ToString.Substring(1, item.ToString.Length - 2))
                    Catch ex As Exception
                    End Try
                Next
            End If
            str_arr = parse_str.Substring(1, txt_len - 2).Split(m_separator)
            Dim idx As Int16
            For idx = 0 To str_arr.Length - 2 Step 2
                If (header_hash.ContainsKey(str_arr(idx))) Then
                    header_hash.Item(str_arr(idx)) = str_arr(idx + 1)
                Else
                    header_hash.Add(str_arr(idx), str_arr(idx + 1))
                End If
            Next
            m_header = header_hash
        Else
            Valid = False
            InValid = "Read Error - Text Header"
            m_header = New Hashtable
            Return
        End If

        txt_Reader.Close()

    End Sub

    Private mLastError As String
    Public ReadOnly Property LastError() As String
        Get
            Return mLastError
        End Get
    End Property

    ' dgp rev 2/18/09 Test the integrity of the actual data bytes
    Public Function DataIntegrity() As Boolean

        Dim DI As New FileInfo(FullSpec)
        Dim check1 = (DI.Length > Pointers(3))
        Dim tot = Header.Item("$TOT")
        Dim check2 = False
        Dim check3 = False
        Dim vbytes = 0
        Dim check4 = True

        If (Header.ContainsKey("$TOT")) Then
            If (Header.ContainsKey("$P1B")) Then
                If (Header.ContainsKey("$PAR")) Then
                    check2 = True
                    Try
                        vbytes = CInt(Header.Item("$TOT")) * CInt(Header.Item("$PAR")) * (CInt(Header.Item("$P1B") / 8))
                    Catch ex As Exception

                    End Try
                    DI = New FileInfo(FullSpec)
                    check3 = (DI.Length > Pointers(2) + vbytes - 1)
                Else
                    mLastError = "Missing $PAR"
                End If
            Else
                mLastError = "Missing $P1B"
            End If
        Else
            mLastError = "Missing $TOT"
        End If
        Return (check1 And check2 And check3 And check4)

    End Function

    ' FCS header information in a hashtable
    Public Property Header() As Hashtable
        Get
            If (m_header Is Nothing) Then Read_Header()
            Return m_header
        End Get
        Set(ByVal Value As Hashtable)
            m_header = Value
        End Set
    End Property

    'dgp rev 12/21/07
    Public Function Read_Subset(ByVal offset As Int32, ByVal count As Int32) As Byte()

        Dim bin_Reader As New BinaryReader(File.Open(FullSpec, FileMode.Open))
        bin_Reader.BaseStream.Seek(Pointers(2) + offset, SeekOrigin.Begin)

        ' Read and verify the data.
        Dim block() As Byte = bin_Reader.ReadBytes(count)
        bin_Reader.Close()

        Return block

    End Function

    Private m_Swap As Boolean = False

    ' Data Block in raw byte format
    Private m_datablock() As Byte
    ' FCS header information in a hashtable
    ' Read in the text header
    Private Sub Read_Data()

        Dim block_len As Integer = Pointers(3) - Pointers(2) + 1
        ' dgp rev 3/5/08 change reader to readonly access to allow protected files
        Dim bin_Reader As New BinaryReader(File.Open(FullSpec, FileMode.Open, FileAccess.Read))
        bin_Reader.BaseStream.Seek(Pointers(2), SeekOrigin.Begin)

        ' Read and verify the data.
        Dim block() As Byte = bin_Reader.ReadBytes(block_len)

        If (block.Length = block_len) Then
            Console.WriteLine("Valid")
            Datablock = block
        Else
            Valid = False
            InValid = "Read Error - Binary Data"
        End If

        bin_Reader.Close()

    End Sub

    ' dgp rev 1/23/09 Restore original pointers
    Private Sub Copy_Pointers()

        For Idx As Int16 = 0 To 3
            m_RevisedPointers(Idx) = Pointers(Idx)
        Next

    End Sub

    ' prep the FCS file instance for writing
    Public Function Write_Prep() As Boolean

        Dim str As String = Separator
        Dim item As DictionaryEntry
        Dim sep As Char = Separator
        ' create the text block from the header hash
        For Each item In Header
            If item.Value Is DBNull.Value Then item.Value = " "
            str = str + CStr(item.Key) + sep + CStr(item.Value) + sep
        Next

        ' update the pointers
        RevisedPointers(1) = str.Length + RevisedPointers(0) - 1
        If (RevisedPointers(1) >= RevisedPointers(2)) Then
            RevisedPointers(2) = Int((RevisedPointers(1) + 1024) / 1024) * 1024
        End If
        m_textblock = str.PadRight(RevisedPointers(2))
        RevisedPointers(3) = Datablock.Length + RevisedPointers(2) - 1
        ' create fixed portion
        Dim Fixed_Str As String
        Fixed_Str = Version.PadRight(10)

        Dim idx As Integer
        For idx = 0 To 3
            Fixed_Str = Fixed_Str + CStr(RevisedPointers(idx)).PadLeft(8)
        Next

        Fixed_Str.PadRight(RevisedPointers(0))
        Dim encoding As New System.Text.ASCIIEncoding

        Fixed = encoding.GetBytes(Fixed_Str)

    End Function

    Private mChkSum() As Byte

    ' dgp rev 4/16/09 
    ' dgp rev 5/4/09 when the data block is load, the checksum is calculated
    Public Property Datablock() As Byte()
        Get
            If (m_datablock Is Nothing) Then Read_Data()
            Return m_datablock
        End Get
        Set(ByVal Value() As Byte)
            If (Not Value Is Nothing) Then
                If (Not m_datablock Is Nothing) Then
                    If (Not m_datablock.Length = Value.Length) Then ReDim m_datablock(Value.Length - 1)
                End If
                mChkSum = md5_obj.ComputeHash(Value)
            End If
            m_datablock = Value
        End Set
    End Property

    ' Required Header fields
    Private m_required() As String
    ' initialize an instance of class
    Public Function init(ByVal filespec As String) As Boolean

        ' make sure file exists
        If (Not File.Exists(filespec)) Then Return False

        Dim f_info As New FileInfo(filespec)

        ' assign properties
        FullSpec = filespec

        FileName = f_info.Name

        Path = f_info.DirectoryName

    End Function

    ' dgp rev 5/9/06 Calculate Checksum
    Public Function ChkSumStr() As String

        ' dgp rev 4/16/09 calculate checksum string from checksum byte array
        ChkSumStr = ""
        If (ChkSum Is Nothing) Then Exit Function
        '        Dim md5_obj As New MD5CryptoServiceProvider
        If (Datablock Is Nothing) Then
            Console.WriteLine("Checksum abort " + FileName)
        Else
            Console.WriteLine("Checksum calculate for " + FileName)
            Dim bite As Byte
            For Each bite In mChkSum
                ChkSumStr = ChkSumStr + (String.Format("{0:X2}", bite))
            Next
        End If

    End Function

    Public ReadOnly Property ChkSum() As Byte()
        Get
            If (mChkSum Is Nothing) Then Dim len = Datablock.Length
            Return mChkSum
        End Get
    End Property
    ' dgp rev 5/9/06 Calculate Checksum
    Public Function Calc_ChkSum() As Byte()

        Return ChkSum

    End Function

    ' dgp rev 7/20/07 save this file
    Private Function Perform_Update() As Boolean

        'Dim sw As New StreamWriter(new_spec, False)
        Dim block As Byte() = Datablock
        Dim encoding As New System.Text.ASCIIEncoding

        Write_Prep()

        Dim fs As New FileStream(FileName, FileMode.Create)
        fs.Write(Fixed, 0, Fixed.Length)
        Dim offset As Int32 = RevisedPointers(0)
        Dim blk_len As Int32 = RevisedPointers(1) - RevisedPointers(0) + 1
        fs.Seek(offset, SeekOrigin.Begin)
        fs.Write(encoding.GetBytes(Text_Block), 0, blk_len)
        offset = RevisedPointers(2)
        blk_len = RevisedPointers(3) - RevisedPointers(2) + 1
        fs.Seek(offset, SeekOrigin.Begin)
        fs.Write(Datablock, 0, blk_len)
        fs.Close()

    End Function

    Private m_Bytes As Int16
    Private mDataSwapped As Boolean = False
    Private mSwapChkSum As Byte()
    Private mOrigChkSum As Byte()

    Private Function CompareByteArrays( _
    ByVal abyt1() As Byte, _
    ByVal abyt2() As Byte _
)
        If abyt1.Length <> abyt2.Length Then
            Return False
        Else
            Dim i As Integer
            For i = 0 To abyt1.Length - 1
                If abyt1(i) <> abyt2(i) Then
                    Return False
                End If
            Next i
        End If
        Return True
    End Function

    ' dgp rev 1/22/09 Swap Bytes
    Public Function SwapBytes() As Boolean

        If (mOrigChkSum Is Nothing) Then mOrigChkSum = ChkSum
        mSwapChkSum = ChkSum
        ' dgp rev 1/22/09 Check header for bit size
        m_Bytes = 4
        If (Header.ContainsKey("$P1B")) Then
            m_Bytes = CInt(Header.Item("$P1B")) / 8
        End If

        Dim idx
        Dim TmpBlock() As Byte
        TmpBlock = Nothing
        Dim sz = Pointers(3) - Pointers(2) + 1
        ReDim TmpBlock(sz - 1)
        If (m_Bytes = 2) Then
            For idx = 0 To sz - 2 Step 2
                TmpBlock(idx) = Datablock(idx + 1)
                TmpBlock(idx + 1) = Datablock(idx)
            Next
        Else
            For idx = 0 To sz - 4 Step 4
                TmpBlock(idx) = Datablock(idx + 3)
                TmpBlock(idx + 1) = Datablock(idx + 2)
                TmpBlock(idx + 2) = Datablock(idx + 1)
                TmpBlock(idx + 3) = Datablock(idx)
            Next
        End If
        Datablock = TmpBlock
        mDataSwapped = (Not Me.CompareByteArrays(mOrigChkSum, mChkSum))
        Return mDataSwapped

    End Function

    ' dgp rev 7/20/07 save this file
    Private Function Perform_Save(ByVal Filename As String) As Boolean

        'Dim sw As New StreamWriter(new_spec, False)
        Dim block As Byte() = Datablock
        Dim encoding As New System.Text.ASCIIEncoding

        Write_Prep()

        Dim fs As New FileStream(Filename, FileMode.Create)
        fs.Write(Fixed, 0, Fixed.Length)
        Dim offset As Int32 = RevisedPointers(0)
        Dim blk_len As Int32 = RevisedPointers(1) - RevisedPointers(0) + 1
        fs.Seek(offset, SeekOrigin.Begin)
        fs.Write(encoding.GetBytes(Text_Block), 0, blk_len)
        offset = RevisedPointers(2)
        blk_len = RevisedPointers(3) - RevisedPointers(2) + 1
        fs.Seek(offset, SeekOrigin.Begin)
        fs.Write(Datablock, 0, blk_len)
        fs.Close()
        Datablock = Nothing

    End Function

    ' dgp rev 1/22/09 Destination path for writing FCS file
    Private m_Destination As String
    Public Property Destination() As String
        Get
            Return m_Destination
        End Get
        Set(ByVal value As String)
            If (System.IO.Directory.Exists(value)) Then m_Destination = value
        End Set
    End Property
    Private mSwapByteFlag As Boolean = False
    Public Property SwapByteFlag() As Boolean
        Get
            Return mSwapByteFlag
        End Get
        Set(ByVal value As Boolean)
            mSwapByteFlag = value
        End Set
    End Property

    ' dgp rev 1/15/09 save this file using binary format
    Private Function Binary_Save(ByVal Filename As String) As Boolean

        'Dim sw As New StreamWriter(new_spec, False)
        Dim pre As Byte() = {0, 0}

        Dim encoding As New System.Text.ASCIIEncoding

        Copy_Pointers()

        Write_Prep()

        If (Me.SwapByteFlag) Then Me.SwapBytes()

        m_textblock = "@"
        Dim item
        For Each item In Header
            If item.value Is DBNull.Value Then item.value = " "
            m_textblock = m_textblock + CStr(item.key) + "@" + CStr(item.value) + "@"
        Next

        Dim fs As New FileStream(Filename, FileMode.Create)
        fs.Write(Fixed, 0, Fixed.Length)
        Dim offset As Int32 = RevisedPointers(0)
        Dim blk_len As Int32 = RevisedPointers(1) - RevisedPointers(0) + 1
        Dim buf_len As Int32 = RevisedPointers(2) - RevisedPointers(1) + 1
        Dim buf(buf_len) As Byte
        fs.Seek(offset, SeekOrigin.Begin)
        fs.Write(encoding.GetBytes(m_textblock), 0, blk_len)
        fs.Seek(RevisedPointers(1) + 1, SeekOrigin.Begin)
        fs.Write(buf, 0, buf_len)
        offset = RevisedPointers(2)
        blk_len = RevisedPointers(3) - RevisedPointers(2) + 1
        fs.Close()

        Dim bs As New BinaryWriter(File.OpenWrite(Filename))
        '       mr.Write(Datablock, bytsiz, Datablock.Length)

        Dim pos
        pos = bs.Seek(RevisedPointers(2), SeekOrigin.Begin)
        bs.Write(Datablock)
        bs.Close()

        Datablock = Nothing

    End Function

    ' dgp rev 7/20/07 save this file
    Public Function Save_File(ByVal new_spec As String) As Boolean

        Binary_Save(new_spec)

    End Function

    ' dgp rev 7/20/07 save as original file spec 
    Public Function Save_File() As Boolean

        Binary_Save(System.IO.Path.Combine(Me.Destination, Me.FileName))

    End Function

    ' dgp rev 1/22/09 Toggle the swap flag
    Public Function Tog_Swap() As Boolean

        m_Swap = (Not m_Swap)
        Return m_Swap

    End Function

    Private md5_obj As MD5CryptoServiceProvider

    ' create a new file object
    Public Sub New(ByVal filespec As String)

        Datablock = Nothing
        m_Swap = False

        ' make sure file exists
        If (Not File.Exists(filespec)) Then
            Valid = False
            InValid = "No such file"
            Return
        Else
            Valid = True
            m_Destination = System.IO.Path.GetDirectoryName(filespec)
        End If
        md5_obj = New MD5CryptoServiceProvider

        ' retrieve file information
        Dim f_info As New FileInfo(filespec)

        ' assign properties
        FullSpec = filespec

        FileName = f_info.Name

        Path = f_info.DirectoryName

        ' validate length of file
        If (f_info.Length < m_fixed_size) Then
            Valid = False
            InValid = "File length is only " + CStr(f_info.Length)
            Return
        End If

        ' open a stream
        Dim sr As StreamReader = Nothing

        Try
            sr = New StreamReader(FullSpec)
            ' Do something with sw
        Catch
            InValid = "Open Error - " + Err.Description
            Return
        End Try

        ' read fixed header area
        Dim fixed As String
        Dim char_arr(m_fixed_size) As Char

        If (sr.ReadBlock(char_arr, 0, m_fixed_size) < m_fixed_size) Then
            Valid = False
            InValid = "Read Error - fixed header"
            If (Not m_keep_open) Then sr.Close()
            Return
        End If

        fixed = char_arr
        ' Is the FCS header valid?
        If (fixed.ToUpper.IndexOf("FCS") < 0) Then
            Valid = False
            InValid = "Non FCS file"
            If (Not m_keep_open) Then sr.Close()
            Return
        End If

        Version = fixed.Trim(" ").Split(" ")(0)
        Dim idx As Int16
        Dim offset As Int16 = 10
        ' read in the pointers
        For idx = 0 To 3
            Pointers(idx) = CInt(fixed.Substring(offset, 8))
            offset += 8
        Next
        Copy_Pointers()
        If (Not m_keep_open) Then sr.Close()

    End Sub
End Class
