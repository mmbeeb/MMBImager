Attribute VB_Name = "DiskTableOps"
' Disk Table operations
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

#Const MMBImager = True

' Disk table entry
Public Type disktableentry_type
    DiskNo As Integer
    ValidDisk As Boolean
    Formatted As Boolean
    ReadOnly As Boolean
    DiskTitle As String
End Type

' Disk table
Public Type disktable_type
    Disk(0 To MaxDisks - 1) As disktableentry_type
    ValidDiskCount As Integer
    FormattedDiskCount As Integer
    ImageName As String
End Type

Public Function ReadDiskTable(strFile As String) As disktable_type
    ' Read disk table from image and associated info
On Error GoTo err_
    Dim t As disktable_type
    Dim f As Long
    Dim b() As Byte
    Dim d As Byte
    Dim x As Integer
    Dim y As Integer
    Dim o As Long
    
    t.ImageName = strFile
    
    ReDim b(0 To DiskTableSize - 1)
    
    f = FreeFile
    Open t.ImageName For Binary Access Read As f
    Get f, 1, b

    For x = 0 To MaxDisks - 1
        o = (x + 1) * 16
        d = b(o + 15)
        
        With t.Disk(x)
            .DiskNo = x
            .ValidDisk = False
            .Formatted = False
            .DiskTitle = ""
                
            If d = DiskReadOnly Or d = DiskReadWrite Then
                .ValidDisk = True
                .Formatted = True
                .ReadOnly = d = DiskReadOnly
            ElseIf d = DiskUnformatted Then
                .ValidDisk = True
            End If
        
            If .ValidDisk Then
                ' Read disk title
                For y = 0 To 11
                    d = b(o + y)
                    If d >= 32 Then
                        .DiskTitle = .DiskTitle & Chr(d)
                    Else
                        Exit For
                    End If
                Next
                
                t.ValidDiskCount = t.ValidDiskCount + 1
                If .Formatted Then
                    t.FormattedDiskCount = t.FormattedDiskCount + 1
                End If
            End If
        End With
    Next
    
    ReadDiskTable = t
    
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Reading Disk Table"
    Resume exit_
End Function

#If Not MMBImager Then
Public Sub UpdateDiskTable(t As disktable_type, _
                            intDiskNo As Integer, _
                            strNewTitle As String)
    ' Update title in disk table for disk
On Error GoTo err_
    Dim b(0 To 11) As Byte
    Dim x As Integer
    Dim f As Long
    Dim o As Long
    
    f = FreeFile
    Open t.ImageName For Binary Access Write As f
    
    o = (intDiskNo + 1) * 16
    
    With t.Disk(intDiskNo)
        .DiskTitle = Left(strNewTitle, 12)
        For x = 0 To Len(.DiskTitle) - 1
            b(x) = Asc(Mid(.DiskTitle, x + 1, 1))
        Next
    End With
    
    ' Write to disk table
    Put f, o + 1, b
    
exit_:
On Error Resume Next
    Close f
    Exit Sub
err_:
    eBox "Modifying Disk Table"
    Resume exit_
End Sub
#End If

#If MMBImager Then
Public Sub ModifyDiskTable(t As disktable_type, _
                            intDiskNo As Integer, _
                            boolRefreshTitle As Boolean)
    ' Update disk table for disk
On Error GoTo err_
    Dim b(0 To 15) As Byte
    Dim c() As Byte
    Dim f As Long
    Dim o As Long
    Dim y As Integer
    
    f = FreeFile
    Open t.ImageName For Binary Access Read Write As f
    
    If boolRefreshTitle Then
        ' Read disk catalogue from image
        ReDim c(0 To DiskCatalogueSize - 1)
        
        o = Disk1Offset + intDiskNo * DiskSize + 1
        Get f, o, c
    End If
    
    o = (intDiskNo + 1) * 16
    Get f, o + 1, b
    
    With t.Disk(intDiskNo)
        ' Update status byte
        If .ValidDisk Then
            If Not .Formatted Then
                b(15) = DiskUnformatted
            ElseIf .ReadOnly Then
                b(15) = DiskReadOnly
            Else
                b(15) = DiskReadWrite
            End If
        Else
            b(15) = DiskInvalid
        End If
        
        If boolRefreshTitle Then
            .DiskTitle = ReadDiskTitle(c)
            
            ' Write title to disk table
            For y = 0 To 11
                If y < Len(.DiskTitle) Then
                    b(y) = Asc(Mid(.DiskTitle, y + 1, 1))
                Else
                    b(y) = 0
                End If
            Next
        End If
    End With
    
    ' Write to disk table
    Put f, o + 1, b
    
exit_:
On Error Resume Next
    Close f
    Exit Sub
err_:
    eBox "Modifying Disk Table"
    Resume exit_
End Sub

Public Function NewDisk(t As disktable_type, _
                Optional intStartDiskNo As Integer = 0) As Integer
    ' Return no. of new disk (first unformatted disk)
    ' or -1 if no free disks
On Error GoTo err_
    Dim x As Integer
    
    NewDisk = -1
    
    x = intStartDiskNo
    If x < 0 Then x = 0
    
    Do While x < t.ValidDiskCount And NewDisk < 0
        With t.Disk(x)
            If Not .Formatted Then
                NewDisk = x
                .Formatted = True
                .ReadOnly = True
                t.FormattedDiskCount = t.FormattedDiskCount + 1
                ModifyDiskTable t, x, True
            End If
        End With
        x = x + 1
    Loop
    'Debug.Print "NewDisk = "; NewDisk
    
exit_:
On Error Resume Next
    Exit Function
err_:
    eBox "New Disk"
    Resume exit_
End Function

Public Sub RebuildDiskTable(ByRef t As disktable_type)
    ' Rebuild disk table
On Error GoTo err_
    Dim f As Long
    Dim b() As Byte
    Dim c() As Byte
    Dim d As Byte
    Dim x As Integer
    Dim y As Integer
    Dim s As Long
    Dim z As Long
    Dim o As Long
    Dim zd As Integer
    Dim l As String
    
    ' Check image size and calculate
    ' no. of disks (zd)
    
    z = FileLen(t.ImageName)
    
    If z < DiskTableSize Then
        z = 0
    Else
        z = z - DiskTableSize
        zd = (z / DiskSize)
        If zd * DiskSize <> z Then
            z = 0
        End If
    End If
    
    If z = 0 Then
        xBox "Invalid MMB file size!"
        Exit Sub
    End If
    
    ReDim b(0 To DiskTableSize - 1)
    ReDim c(0 To DiskCatalogueSize - 1)
    
    ' Read disk table
    f = FreeFile
    Open t.ImageName For Binary Access Read Write As f
    Get f, 1, b
    
    For x = 0 To MaxDisks - 1
        With t.Disk(x)
            o = (x + 1) * 16
        
            ' Validate status
            .ValidDisk = False
            .Formatted = False
            .ReadOnly = True
        
            If x < zd Then
                d = b(o + 15)
        
                If d = DiskReadWrite Then
                    .Formatted = True
                    .ReadOnly = False
                ElseIf d = DiskReadOnly Then
                    .Formatted = True
                End If
                
                .ValidDisk = True
            End If
            
            d = DiskInvalid
            
            If .ValidDisk Then
                If .Formatted And .ReadOnly Then
                    d = DiskReadOnly
                ElseIf .Formatted Then
                    d = DiskReadWrite
                Else
                    d = DiskUnformatted
                End If
            End If
            
            b(o + 15) = d
            
            If .ValidDisk Then
                s = Disk1Offset + x * DiskSize + 1
                Get f, s, c
                
                .DiskTitle = ReadDiskTitle(c)
                
                ' Write title to disk table
                For y = 0 To 11
                    If y < Len(.DiskTitle) Then
                        b(o + y) = Asc(Mid(.DiskTitle, y + 1, 1))
                    Else
                        b(o + y) = 0
                    End If
                Next
            End If
        End With
    Next
    
    ' Write disk table
    Put f, 1, b
    
exit_:
On Error Resume Next
    Close f
    Exit Sub
err_:
    eBox "Rebuild Disk Table"
    Resume exit_
End Sub

Public Function ReadDiskTitle(cat() As Byte) As String
    ' Read title from disk catalogue
    Dim l As String
    Dim y As Long
    Dim d As Byte
    
    l = ""
    y = 0
    Do
        If y > 7 Then
            d = cat(y + &HF8)
        Else
            d = cat(y)
        End If
        
        If d > 0 Then
            l = l & Chr(d)
        End If
        
        y = y + 1
    Loop Until y = 12 Or d = 0
    
    ReadDiskTitle = Trim(l)
End Function
#End If
