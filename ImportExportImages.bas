Attribute VB_Name = "ImportExportImages"
' Import / Export .SSD disk image
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

Public Function DiskPtr(intDiskNo As Integer, _
                Optional intSector As Integer = 0) As Long
    ' Return pointer to disk sector in image file
    ' Remember, the first byte is at position '1' in the image file
    DiskPtr = Disk1Offset + intDiskNo * _
                    DiskSize + intSector * SecSize + 1
End Function

Public Function ExportDFSDisk(lngSrcFileHandle As Long, _
                    d As disktableentry_type, _
                    boolFullSize As Boolean) As String
    ' Extract DFS disk image
On Error GoTo err_
    Dim lngTgt As Long
    Dim strName As String
    Dim strPath As String
    Dim lngLen As Long
    Dim lngDiskPtr As Long
    Dim bytData() As Byte
    Dim bytTemp As Byte
    Dim x As Integer
    
    Debug.Print "Extract disk "; d.DiskNo, _
            d.DiskTitle, d.ValidDisk, d.Formatted
            
    ExportDFSDisk = ""
    
    If d.ValidDisk And d.Formatted Then
        ' Target file name:
        strName = Trim(d.DiskTitle)
        
        If strName = "" Then
            strName = "Disk_" & d.DiskNo & "_Untitled"
        End If
        
        strName = strName & ".ssd"
    
        lngDiskPtr = DiskPtr(d.DiskNo)
        
        ' Size of target
        If boolFullSize Then
            lngLen = DiskSize
        Else
            ' Calc minimum size
            lngLen = DiskMinSize(lngSrcFileHandle, lngDiskPtr) * SecSize
        End If
        
        'Debug.Print "EX DISK "; strName; "  LEN="; lngLen
        
        strPath = TempFolder(strName)
        
        ReDim bytData(1 To lngLen)
        
        ' Kill temporary file if it already exists
        If Dir(strPath) <> "" Then
            Kill strPath
        End If
            
        lngTgt = FreeFile
        Open strPath For Binary Access Write As lngTgt
                
        Get lngSrcFileHandle, lngDiskPtr, bytData
                    
        Put lngTgt, , bytData
    
        ExportDFSDisk = strPath
    End If
    
exit_:
On Error Resume Next
    Close lngTgt
    Exit Function
err_:
    eBox "Export Disk"
    Resume exit_
End Function

Public Function ImportDFSImage(intSpecificDiskNo As Integer, _
                                DTable As disktable_type, _
                                strDosFilename As String) As Integer
    ' Import dfs images
    ' Returns disk no. or -1 on error
On Error GoTo err_
    Dim strExt As String
    Dim lngTgt As Long
    Dim lngSrc As Long
    Dim lngLen As Long
    Dim intDiskNo As Integer
    Dim lngDiskPtr As Long
    Dim bytTemp As Byte
    Dim bytData() As Byte
    Dim boolSSD As Boolean
    Dim intSectors As Integer
    Dim intMinSectors As Integer
    Dim intMaxSectors As Integer
    Dim x As Integer
    
    ImportDFSImage = -1
    
    ' Get type of image, and calc max sectors
    strExt = LCase(Right(strDosFilename, 4))
    boolSSD = strExt = ".ssd" Or strExt = ".img"
    
    intMinSectors = 2
    intMaxSectors = DiskSectors
    
    If boolSSD Then
        lngLen = FileLen(strDosFilename)
        intSectors = lngLen \ SecSize
        
        ' Check file size in range and
        ' a multiple of sector size
        If intSectors >= intMinSectors And _
                intSectors <= intMaxSectors And _
                intSectors * SecSize = lngLen Then

            'Debug.Print " >> secs="; intSectors
            
            ' Get disk no, returns -1 if none available
            intDiskNo = NewDisk(DTable, intSpecificDiskNo)
            
            If intDiskNo >= 0 Then
                'Debug.Print " >> disk "; intDiskNo
                
                ReDim bytData(1 To lngLen)
                
                ' Open Target (.mmb)
                lngTgt = FreeFile
                Open DTable.ImageName For Binary Access Write As lngTgt
                
                ' Open Source (.ssd or .dsd)
                lngSrc = FreeFile
                Open strDosFilename For Binary Access Read As lngSrc
                
                lngDiskPtr = DiskPtr(intDiskNo)
                
                Get lngSrc, , bytData
                    
                Put lngTgt, lngDiskPtr, bytData
                
                Debug.Print "Imported image "; strDosFilename, intDiskNo
                
                ModifyDiskTable DTable, intDiskNo, True ' Refresh title
                
                ImportDFSImage = intDiskNo
            Else
                xBox "No more free disks!"
            End If
        End If
    End If
exit_:
On Error Resume Next
    Close lngTgt
    Close lngSrc
    Exit Function
err_:
    eBox "Import Image"
    Resume exit_
End Function

Private Function DiskMinSize(lngSrcFileHandle As Long, _
                            lngDiskPtr As Long) As Integer
    ' Calc. min. disc size
    Dim cat(0 To DiskCatalogueSize - 1) As Byte
    Dim intFiles As Integer
    Dim intSize As Integer
    Dim y As Integer
    Dim o As Integer
    Dim z As Long
    Dim mixedbyte As Byte
    
    ' Read disk catalogue
    Get lngSrcFileHandle, lngDiskPtr, cat
        
    intFiles = cat(&H105) / 8
    intSize = 2
    
    For y = 1 To intFiles
        ' Calc last sector used by file + 1
        o = y * 8 + &H100
    
        mixedbyte = cat(o + 6)
        
        z = (mixedbyte And &H3) * &H100 + CLng(cat(o + 7)) + _
            ((mixedbyte And &H30) \ &H10) * &H100 + CLng(cat(o + 5))
            
        If cat(o + 4) > 0 Then z = z + 1
        
        If z > intSize Then intSize = z
    Next
    
    DiskMinSize = intSize
End Function
