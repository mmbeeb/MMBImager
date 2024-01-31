Attribute VB_Name = "DiskOps"
' General Disk Operations
' Written by Martin Mather 2005

Option Explicit

' Catalogue file entry
Public Type cataloguefile_Type
    Name As String
    Directory As String
    FullName As String ' Dir + . + Name
    Locked As Boolean
    Load As Long
    Exec As Long
    Length As Long
    StartSector As Integer
    SectorsUsed As Integer
End Type

' Disk catalogue
Public Type catalogue_type
    ImageFile As String
    DiskNo As Integer
    DiskTitle As String
    CycleNo As Byte
    Option As Byte
    Sectors As Integer
    SectorsUsed As Integer
    LastSector As Integer
    FileCount As Byte
    Files(1 To CatalogueMaxFiles) As cataloguefile_Type
    Modified As Boolean
End Type

Public Function ReadDiskCatalogue(strFile As String, _
            DiskNo As Integer) As catalogue_type
    ' Read Disk Catalogue
On Error GoTo err_
    Dim cat(0 To 511) As Byte
    Dim c As catalogue_type
    Dim f As Long
    Dim x As Integer
    Dim y As Integer
    Dim o As Integer
    Dim b As Byte
    Dim s As Long
    Dim mixedbyte As Byte
    
    Debug.Print "ReadDiskCatalogue: "; DiskNo
    
    ' Read disk catalogue
    f = FreeFile
    Open strFile For Binary Access Read As f
    s = DiskPtr(DiskNo)
    Get f, s, cat
    Close f
    
    With c
        .ImageFile = strFile
        .DiskNo = DiskNo
        .Modified = False
        
        ' Read disk title (Chr 0 terminated)
        x = 0
        Do
            If x > 7 Then b = cat(x + &HF8) Else b = cat(x)
            If b > 0 Then
                .DiskTitle = .DiskTitle & Chr(b)
            End If
            x = x + 1
        Loop Until x = 11 Or b = 0
        
        .Option = (cat(&H106) And &HF0) \ &H10
        .Sectors = (cat(&H106) And &H3) * &H100 + cat(&H107)
        .CycleNo = BCDtoBin(cat(&H104))
        .FileCount = cat(&H105) / 8
        
        For y = 1 To .FileCount
            With .Files(y)
                o = y * 8
                
                ' Filename (padded with spaces)
                For x = 0 To 6
                    .Name = .Name & Chr(cat(o + x))
                Next
                .Name = RTrim(.Name)
                .Directory = Chr(cat(o + 7) And &H7F)
                .FullName = .Directory & "." & .Name
                .Locked = cat(o + 7) >= &H80
        
                o = o + &H100
                mixedbyte = cat(o + 6)
                
                .Load = CLng(cat(o + 1)) * &H100 + CLng(cat(o))
                If (mixedbyte And &HC) > 0 Then .Load = .Load + &HFF0000
 
                .Exec = CLng(cat(o + 3)) * &H100 + CLng(cat(o + 2))
                If (mixedbyte And &HC0) > 0 Then .Exec = .Exec + &HFF0000

                .Length = ((mixedbyte And &H30) \ &H10) * &H10000 + _
                    CLng(cat(o + 5)) * &H100 + CLng(cat(o + 4))

                .StartSector = (mixedbyte And &H3) * &H100 + CLng(cat(o + 7))

                .SectorsUsed = ((mixedbyte And &H30) \ &H10) * &H100 + _
                    CLng(cat(o + 5))
                If cat(o + 4) > 0 Then .SectorsUsed = .SectorsUsed + 1
                
                c.SectorsUsed = c.SectorsUsed + .SectorsUsed
            End With
        Next
    End With
    
    ReadDiskCatalogue = c
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Reading Catalogue"
    Resume exit_
End Function

Public Sub RefreshDiskCatalogue(ByRef Catalogue As catalogue_type)
    ' Refresh Disk Catalogue
    Catalogue = ReadDiskCatalogue(Catalogue.ImageFile, Catalogue.DiskNo)
End Sub

Public Sub SaveDiskCatalogue(Catalogue As catalogue_type)
    ' Save Disk Catalogue
On Error GoTo err_
    Dim cat(0 To 511) As Byte
    Dim f As Long
    Dim x As Integer
    Dim y As Integer
    Dim o As Integer
    Dim b As Byte
    Dim s As Long
    Dim mixedbyte As Byte
    
    Debug.Print "SaveDiskCatalogue: "; Catalogue.DiskNo
    With Catalogue
        ' Increment cycle no.
        .CycleNo = (.CycleNo + 1) Mod 100
        
        ' Check & fix file list
        For y = 1 To .FileCount
            Debug.Print "sc "; y, .Files(y).FullName
        Next
        
        ' Disk title (Chr 0 terminated)
        For x = 0 To 10
            If x + 1 > Len(.DiskTitle) Then
                b = 0
            Else
                b = Asc(Mid(.DiskTitle, x + 1, 1))
            End If
            
            If x > 7 Then cat(x + &HF8) = b Else cat(x) = b
        Next
        
        cat(&H106) = (.Option * &H10) Or (.Sectors \ &H100)
        cat(&H107) = .Sectors And &HFF
        cat(&H104) = BintoBCD(.CycleNo)
        cat(&H105) = .FileCount * 8
        
        For y = 1 To .FileCount
            With .Files(y)
                o = y * 8
                
                ' Filename (padded with spaces)
                For x = 0 To 6
                    cat(o + x) = Asc(Mid(.Name & String(7, " "), x + 1, 1))
                Next
                cat(o + 7) = Asc(.Directory) Or IIf(.Locked, &H80, 0)
     
                o = o + &H100
                mixedbyte = 0
                
                cat(o) = .Load And &HFF
                cat(o + 1) = (.Load \ &H100) And &HFF
                If .Load >= &H10000 Then mixedbyte = &HC
 
                cat(o + 2) = .Exec And &HFF
                cat(o + 3) = (.Exec \ &H100) And &HFF
                If .Exec >= &H10000 Then mixedbyte = mixedbyte Or &HC0

                cat(o + 4) = .Length And &HFF
                cat(o + 5) = (.Length \ &H100) And &HFF
                mixedbyte = mixedbyte Or (.Length \ &H1000 And &H30)
                
                cat(o + 7) = .StartSector And &HFF
                mixedbyte = mixedbyte Or (.StartSector \ &H100 And 3)
                
                cat(o + 6) = mixedbyte
            End With
        Next
    End With
    
    ' Write disk catalogue
    f = FreeFile
    Open Catalogue.ImageFile For Binary Access Write As f
    s = DiskPtr(Catalogue.DiskNo)
    Put f, s, cat
    Close f
        
    Catalogue.Modified = False
exit_:
On Error Resume Next
    Close f
    Exit Sub
err_:
    eBox "Writing Catalogue"
    Resume exit_
End Sub

Private Function GetFileIndex(ByRef cat As catalogue_type, _
                strFileName As String) As Byte
    ' Return file index
    Dim y As Integer
    GetFileIndex = 0
    With cat
        For y = 1 To .FileCount
            If .Files(y).FullName = strFileName Then
                GetFileIndex = y
                Exit Function
            End If
        Next
    End With
End Function

Public Function ExtractDFSFile(lngSrcFileHandle As Long, _
        intFileNo As Integer, cat As catalogue_type) As String
    ' Saves dfs file (and .inf) from image to temporary folder
    ' Returns pathname, or "" if error occurs
On Error GoTo err_
    Dim f As Long
    Dim strName As String
    Dim strPath As String
    Dim strInf As String
    Dim lngPtr As Long
    Dim bytData() As Byte
    
    ExtractDFSFile = ""
    With cat.Files(intFileNo)
        strName = .Directory & "." & .Name
        strPath = TempFolder(strName)
        ReDim bytData(1 To .Length)
        lngPtr = DiskPtr(cat.DiskNo, .StartSector)
        Get lngSrcFileHandle, lngPtr, bytData
        
        f = FreeFile
        Open strPath For Binary Access Write As f
        Put f, , bytData
        Close f
    End With
    
    If WriteInf(strPath, cat.Files(intFileNo), bytData) Then
        ExtractDFSFile = strPath
    End If
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Extract file"
    Resume exit_
End Function

Public Function ImportDFSFile(cat As catalogue_type, _
                strDosFilename As String) As Boolean
    ' Import file in to disk image
On Error GoTo err_
    Dim strExt As String
    Dim f As Long
    Dim file As cataloguefile_Type
    Dim l As Long
    
    ImportDFSFile = True
    l = FileLen(strDosFilename)
    If l > 0 And l <= 200 * KB Then
        strExt = LCase(Right(strDosFilename, 4))
        If strExt <> ".inf" Then ' Ignore .inf files
            If strExt = ".img" Or strExt = ".ssd" Then
                ' Pucrunch image
                Debug.Print "Pucrunch: "; strDosFilename
                Exit Function
            End If
            
            ' Read .inf file (must exist)
            If ReadInf(strDosFilename, file) Then
                ' read file
                Dim bytData() As Byte
                ReDim bytData(0 To file.SectorsUsed * SecSize - 1)
                f = FreeFile
                Open strDosFilename For Binary Access Read As f
                Get f, , bytData
                Close f
            
                ImportDFSFile = AddFile(cat, file, bytData)
            End If
        End If
    End If
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    ImportDFSFile = False
    eBox "Import file"
    Resume exit_
End Function

Private Function WriteInf(strDosFilename As String, _
                file As cataloguefile_Type, bytData() As Byte) As Boolean
    ' Create .inf file
On Error GoTo err_
    Dim strInf As String
    Dim f As Long
    
    WriteInf = False
    With file
        strInf = Left(.FullName & String(8, " "), 10) _
                & HexN(.Load, 6, " ") & " " _
                & HexN(.Exec, 6, " ")
        If .Locked Then strInf = strInf & " Locked"
        strInf = strInf & " CRC= " & _
                Hex(CalcCRC(bytData, .Length))
    End With
    
    f = FreeFile
    Open strDosFilename & ".inf" For Binary Access Write As f
    Put f, , strInf
    Close f
    WriteInf = True
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Write .inf"
    Resume exit_
End Function

Private Function ReadInf(strDosFilename As String, _
                file As cataloguefile_Type) As Boolean
    ' Read .inf file
    ' Returns True if successful
On Error GoTo err_
    Dim f As Long
    Dim strInf As String
    
    ReadInf = False
    With file
        .Length = FileLen(strDosFilename)
        If .Length > 0 Then
            Debug.Print "DFSFile: "; strDosFilename, .Length
                
            If Dir(strDosFilename & ".inf", vbNormal) <> "" Then
                f = FreeFile
                Open strDosFilename & ".inf" For Input As f
                Input #f, strInf
                Close f
                Debug.Print "Inf: "; strInf
                
                .FullName = Parse(strInf)
                .Directory = Left(.FullName, 1)
                .Name = Mid(.FullName, 3)
                
                .Load = CLng("&H" & Parse(strInf))
                .Exec = CLng("&H" & Parse(strInf))
                .Locked = InStr(UCase(Mid(strInf, 25)), "LOCKED") > 0
                    
                .SectorsUsed = (.Length \ &H100)
                If (.Length And &HFF) > 0 Then .SectorsUsed = .SectorsUsed + 1
                ReadInf = .Name <> "" And .Directory <> ""
            End If
        End If
    End With
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Read .inf"
    Resume exit_
End Function

Private Function AddFile(cat As catalogue_type, _
                    file As cataloguefile_Type, bytData() As Byte) As Boolean
    ' Add file to DFS disk
    ' Returns true if successful
    Dim f As Long
    Dim y As Integer
    Dim p As Integer
    Dim o As Boolean
    Dim lngPtr As Long
    
    AddFile = False
    
    RefreshDiskCatalogue cat
    
    ' Catalogue full?
    y = GetFileIndex(cat, file.FullName)
    If y = 0 And cat.FileCount = 31 Then
        MsgBox "Catalogue full!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    
    ' Overwrite existing file?
    If y > 0 Then
        o = MsgBox("Overwrite " & file.FullName, _
                    vbYesNo Or vbExclamation) = vbYes
        If o Then
            DeleteDFSFile cat, file.FullName
        End If
    Else
        o = True
    End If
    
    If o Then
        With file
            ' Room on disk?
            If .SectorsUsed > (cat.Sectors - cat.SectorsUsed) Then
                MsgBox "No room on disk!", vbExclamation Or vbOKOnly
            Else
                .StartSector = GetDiskBlock(cat, .SectorsUsed)
                If .StartSector = 0 Then
                    ' There is enough room if the disk is compacted
                    If MsgBox("No block large enough!" & vbNewLine & _
                            "Compact disk?", vbExclamation Or vbYesNo) = vbYes Then
                        CompactDisk cat
                        '.StartSector = cat.LastSector
                    End If
                End If
                
                If .StartSector > 0 Then
                    Debug.Print "FN   "; .FullName
                    Debug.Print "Dir  "; .Directory
                    Debug.Print "N    "; .Name
                    Debug.Print "Load "; Hex$(.Load)
                    Debug.Print "Exec "; Hex$(.Exec)
                    Debug.Print "lock "; .Locked
                    Debug.Print "SU   "; .SectorsUsed
                    Debug.Print "strs "; .StartSector
                    
                    lngPtr = DiskPtr(cat.DiskNo, .StartSector)
                    Debug.Print "@ptr:: "; lngPtr
                    
                    ' write file
                    f = FreeFile
                    Open cat.ImageFile For Binary Access Write As f
                    Put f, lngPtr, bytData
                    Close f
                    
                    ' add to catalogue (must be in 'start sector' desc. order)
                    With cat
                        p = 1
                        For y = 1 To .FileCount
                            If .Files(y).StartSector < file.StartSector Then
                                p = y
                                Exit For
                            Else
                                p = y + 1
                            End If
                        Next
                        Debug.Print "Cat ptr="; p
                        If p <= .FileCount Then ' Insert gap
                            For y = .FileCount To p Step -1
                                .Files(y + 1) = .Files(y)
                            Next
                        End If
                        .Files(p) = file
                        .FileCount = .FileCount + 1
                    End With
                    SaveDiskCatalogue cat
                    AddFile = True
                End If
            End If
        End With
    End If
End Function

Public Function GetDiskBlock(cat As catalogue_type, intSize As Integer) As Integer
    ' Return start sector of smallest free block of minimum size of intSize
    ' If zero returned, no block found
    Dim x As Integer
    Dim s As Integer
    Dim b As Integer
    Dim z As Integer
    
    GetDiskBlock = 0
    With cat
        If .FileCount = 0 Then
            GetDiskBlock = 2
        Else
            s = 2
            z = 0
            For x = .FileCount To 0 Step -1
                If x = 0 Then
                    b = .Sectors - s
                    Debug.Print x, "END OF DISK", Hex$(s), Hex$(.Sectors), b
                Else
                    b = .Files(x).StartSector - s
                    Debug.Print x, .Files(x).FullName, Hex$(s), Hex$(.Files(x).StartSector), b
                End If
                    
                If b >= intSize And (b < z Or z = 0) Then
                    GetDiskBlock = s
                    z = b
                End If
                    
                If x > 0 Then
                    s = .Files(x).StartSector + .Files(x).SectorsUsed
                End If
            Next
            
        End If
    End With
    Debug.Print "GDB : "; GetDiskBlock
End Function

Public Sub DeleteDFSFile(cat As catalogue_type, strFileName As String)
    ' Delete DFS file
    ' Assumes catalogue is current and does not write to disk
    Dim y As Integer
    Dim x As Integer
   
    Debug.Print "Delete file: "; strFileName
    With cat
        y = GetFileIndex(cat, strFileName)
        If y > 0 Then
            If y < .FileCount Then
                For x = y To .FileCount - 1
                    .Files(x) = .Files(x + 1)
                Next
            End If
            .FileCount = .FileCount - 1
        End If
    End With
End Sub

Public Function DiskPtr(intDiskNo As Integer, _
                Optional intSector As Integer = 0) As Long
    ' Return pointer to disk sector in image file
    ' Remember, the first byte is at position '1' in the image file
    DiskPtr = DiskTableSize + intDiskNo * DiskSize + intSector * SecSize + 1
End Function

Public Function CompactDisk(cat As catalogue_type) As Boolean
    ' Compact disk
    ' Assumes catalogue is current
    ' Returns true if any files moved
On Error GoTo err_
    Dim y As Integer
    Dim s As Integer
    Dim z As Integer
    Dim bytData() As Byte
    Dim f As Long
    Dim boolUpdateCat As Boolean
    
    Debug.Print "CompactDisk"
    CompactDisk = False
    f = FreeFile
    boolUpdateCat = False
    Open cat.ImageFile For Binary Access Read Write As f
    s = 2
    With cat
        For y = .FileCount To 1 Step -1
            With .Files(y)
                z = .StartSector - s
                Debug.Print y, Hex$(s), Hex$(.StartSector), z
                If z > 0 Then
                    ' Move file
                    ReDim bytData(1 To .Length)
                    Get f, DiskPtr(cat.DiskNo, .StartSector), bytData
                    Put f, DiskPtr(cat.DiskNo, s), bytData
                    Erase bytData
                    .StartSector = s
                    boolUpdateCat = True
                End If
                s = s + .SectorsUsed
            End With
        Next
    End With
    Close f
    If boolUpdateCat Then
        SaveDiskCatalogue cat
    End If
    CompactDisk = boolUpdateCat
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Compact Disk"
    Resume exit_
End Function

Public Function ExtractDFSDisk(lngSrcFileHandle As Long, d As disktableentry_type, _
                boolFullSize As Boolean) As String
    ' Extract DFS disk image
On Error GoTo err_
    Dim f As Long
    Dim strName As String
    Dim strPath As String
    Dim strInf As String
    Dim lngPtr As Long
    Dim bytData() As Byte
    Dim z As Long
    
    Debug.Print "Extract disk "; d.DiskNo, d.DiskTitle, d.ValidDisk, d.Unformatted
    ExtractDFSDisk = ""
    
    If d.ValidDisk And Not d.Unformatted Then
        strName = Trim(d.DiskTitle)
        If strName = "" Then
            strName = "Disk_" & d.DiskNo & "_Untitled.ssd"
        End If
        strName = strName & ".ssd"
        
        If boolFullSize Then
            z = DiskSize
        Else
            z = d.LastSector * SecSize
        End If
        
        strPath = TempFolder(strName)
        ReDim bytData(1 To z)
        
        lngPtr = DiskPtr(d.DiskNo)
        Get lngSrcFileHandle, lngPtr, bytData
            
        f = FreeFile
        Open strPath For Binary Access Write As f
        Put f, , bytData
        Close f
            
        ExtractDFSDisk = strPath
    End If
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Extract Disk"
    Resume exit_
End Function

Public Function ImportDFSImage(intDiskNo As Integer, DTable As disktable_type, _
                                strDosFilename As String) As Boolean
    ' Import dfs files or images
On Error GoTo err_
    Dim strExt As String
    Dim f As Long
    Dim cat As catalogue_type
    Dim l As Long
    Dim ok As Boolean
    Dim bytData() As Byte
    Dim ptr As Long
    Dim d As Integer
    
    ImportDFSImage = False
    l = FileLen(strDosFilename)
    If l > 0 And l <= 200 * KB Then
        strExt = LCase(Right(strDosFilename, 4))
        If strExt = ".ssd" Or strExt = ".img" Then
            ok = intDiskNo < 0
            ' Overwrite existing disk?
            If Not ok Then
                If DTable.Disk(intDiskNo).Unformatted Then
                    ok = True
                Else
                    ok = MsgBox("Overwrite disk " & vbNewLine & intDiskNo & _
                            ": " & DTable.Disk(intDiskNo).DiskTitle, _
                            vbQuestion Or vbYesNo) = vbYes
                End If
            End If
            If ok Then
                If intDiskNo >= 0 Then
                    d = intDiskNo
                Else
                    d = CreateNewDisk(DTable)
                End If
                
                If d >= 0 Then
                    ' read image
                    ReDim bytData(1 To l)
                    f = FreeFile
                    Open strDosFilename For Binary Access Read As f
                    Get f, , bytData
                    Close f
                    
                    ' write image
                    ptr = DiskPtr(d)
                    Open DTable.ImageName For Binary Access Write As f
                    Put f, ptr, bytData
                    Close f
                    
                    Debug.Print "Imported image "; strDosFilename, d
                    ImportDFSImage = True
                End If
            End If
        ElseIf intDiskNo >= 0 Then
            ' Add files to existing disk
            cat = ReadDiskCatalogue(DTable.ImageName, intDiskNo)
            ImportDFSImage = ImportDFSFile(cat, strDosFilename)
        End If
    End If
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Import Image"
    Resume exit_
End Function
