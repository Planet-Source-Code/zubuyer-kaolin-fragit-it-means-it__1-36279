Attribute VB_Name = "modFHand"
''''''''''''''''''''''''''''''''''''''''
'FragIt                                '
'© Copyright 2002 by Muhammad Zubaer   '
'                                      '
'This is a FREEWARE but this code      '
'is not intend to be used commercially.'
'Although you can use it as you like   '
'in your own project but do not resale '
'it or destroy the original author's   '
'name. If you use this code in your    '
'project than it would be nice to give '
'me some cradits. I've worked hard on  '
'it.
'                                      '
'Warning: There is no warranty provided'
'so use it in your own risk. The author'
'is not responsible for any damage     '
'caused by this code.                  '
'                                      '
'Mail me at the following address if   '
'you have any questions or made any    '
'enhancement.                          '
'lifeforcez@hotmail.com                '
''''''''''''''''''''''''''''''''''''''''
    
    'Type declaration for ChunkSize variable
    Type ChunkSize
        S12000 As String * 12000
        S6000 As String * 6000
        S3000 As String * 3000
        S1500 As String * 1500
        S500 As String * 500
        S100 As String * 100
        S25 As String * 25
        S5 As String * 5
        S1 As String * 1
    End Type
    
    'Declare the variable Bytes as of ChunkSize type
    Dim Bytes As ChunkSize
    
Function SplitFile(FileName As String, SegmentSize As Long, Optional NumOfSegments As Integer) As Integer

    Dim SourceBytes As Long
    Dim SourceFile As String
    Dim DestinationFile As String
    Dim SegmentNumber As Integer
    Dim BytesDone As Long
    Dim FPath As String
    Dim FName As String
    Dim FNameNoExt As String
    Dim ErrorCode As Integer
    
    On Error GoTo ErrorHandler
    
    'Make sure the file exists
    If FileName = "" Or Dir(FileName) = "" Then
        ErrorCode = 1
        GoTo ErrorHandler
    End If
    
    'Ensure that the segment size is valid
    If SegmentSize = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
    
    'Retrieve the path name where file exists
    Do
        i = i + 1
        'Find the first occurance of the "\" in the FileName string from the right
        j = InStr(Len(FileName) - i, FileName, "\", vbTextCompare)
    Loop Until j > 0
    
    'Extract the file name
    FName = Right$(FileName, Len(FileName) - j)
    
    'Extract the path name
    FPath = Left$(FileName, j)
    
    'Find the file name without extension
    'Find the first occurance of the '.' in the FileName string
    j = InStr(1, FName, ".", vbTextCompare)
    If j = 0 Then   'File name does not contain a '.' character
        FNameNoExt = FName
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(FName, j - 1)
    End If
    
    'Get total number or bytes in the source file
    SourceBytes = FileLen(FileName)
    
    'Ensure that the resultant file segments will not exceed 999 segments
    'because otherwise we will have incorrect file extensions
    If SourceBytes / SegmentSize >= 1000 Then
        ErrorCode = 3
        GoTo ErrorHandler
    End If

    'Open the source file for binary read
    Open FileName For Binary Access Read As #1 Len = 1
    
    Do
        'Increase the number of segments counter by 1
        SegmentNumber = SegmentNumber + 1
        
        'Compose the file name of the new file to be created (file segment)
        DestinationFile = FPath & FName & "." & CStr(Format(SegmentNumber, "000"))
                    
        'Create the new file segment and open it for binary write
        Open DestinationFile For Binary Access Write As #2 Len = 1
        
        'Check whether the remaining bytes to process in the source file are
        'less than Segment bytes
        If SourceBytes - BytesDone < SegmentSize Then
            RemainingBytes = SourceBytes - BytesDone
        Else
            RemainingBytes = SegmentSize
        End If
       
       'Read bytes from the source file and write them to the destination file (the current segment file)
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
       Do
            
            Select Case RemainingBytes
                Case Is >= 12000
                    'Read 12000 bytes of data from the source file
                    Get #1, , Bytes.S12000
                    'Write the bytes to the destination file
                    Put #2, , Bytes.S12000
                    'Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    'Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    'Yield to windows and other processes to do their jobs
                    'Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #1, , Bytes.S6000
                    Put #2, , Bytes.S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #1, , Bytes.S3000
                    Put #2, , Bytes.S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #1, , Bytes.S1500
                    Put #2, , Bytes.S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #1, , Bytes.S500
                    Put #2, , Bytes.S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #1, , Bytes.S100
                    Put #2, , Bytes.S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #1, , Bytes.S25
                    Put #2, , Bytes.S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #1, , Bytes.S5
                    Put #2, , Bytes.S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #1, , Bytes.S1
                    Put #2, , Bytes.S1
                    RemainingBytes = RemainingBytes - 1
                    BytesDone = BytesDone + 1
                    DoEvents
                Case Is = 0
                    'When the loop enters here, the segment bytes are completed.
                    'Close the segment file and exit the loop
                    Close 2
                    DoEvents
                    Exit Do
            End Select
            
            'Update the percent control on the form
            frmMain.spProg (Int((BytesDone / SourceBytes) * 100))
            'Refresh the form and yield to windows
            DoEvents
        Loop
        
    Loop Until BytesDone = SourceBytes
    'Close the source file
    Close 1
    
    'When the code reaches this point, everything went OK.
    'Acknowledge the number of segments, assign the value '0' to the function and exit
    NumOfSegments = SegmentNumber
    SplitFile = 0
    frmMain.spProg (0)
    Exit Function
    
ErrorHandler:
    
    'This is entered only when an error occures
    Select Case ErrorCode
        Case Is = 0 'Unknown error
            Reset   'Close any open files
            SplitFile = 4   'Assign error code 4 to the function
        Case Else 'Assign error code value to the function (1 to 3)
            SplitFile = ErrorCode
    End Select
    
    Exit Function

End Function

Function MergeFiles(SourceFile As String, Optional NumOfSegments As Integer) As Integer

    Dim TotalBytes As Long
    Dim DestinationFile As String
    Dim SegmentFile As String
    Dim SegmentNumber As Integer
    Dim Segments As Integer
    Dim BytesDone As Long
    Dim FPath As String
    Dim FName As String
    Dim FNameNoExt As String
    Dim ErrorCode As Integer
    
    On Error GoTo ErrorHandler
    
    'Make sure the source file name is given and is valid (exists)
    If SourceFile = "" Or Dir(SourceFile) = "" Then
        ErrorCode = 1
        GoTo ErrorHandler
    End If
    
    'Find the number of segments of the split file
    'Retrieve the path name where files exist
    Do
        i = i + 1
        'Find the first occurance of the "\" in the SourceFile string from the right
        j = InStr(Len(SourceFile) - i, SourceFile, "\", vbTextCompare)
    Loop Until j > 0
    
    'Extract the file name
    FName = Right$(SourceFile, Len(SourceFile) - j)
    
    'Extract the path name
    FPath = Left$(SourceFile, j)
    
    'Find the file name without extension
    'Find the first occurance of the '.' in the SourceFile string
    j = InStr(1, StrReverse(FName), ".", vbTextCompare)
    If j = 0 Then   'File name does not contain a '.' character
        FNameNoExt = FName
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(FName, Len(FName) - j)
    End If
    
    'Now find the number of segments of the split file that reside in
    'the same directory where the source file is
    'Also count the total number of bytes in the segments (this will be
    'used for the calculation of the percent done value
    Do
        'Increase the number of segments counter by 1
        Segments = Segments + 1
        
        'Compose the segment file name and check
        SegmentFile = FPath & FNameNoExt & "." & CStr(Format(Segments, "000"))
        If Dir(SegmentFile) = "" Then Exit Do
        TotalBytes = TotalBytes + FileLen(SegmentFile)
    Loop
    
    Segments = Segments - 1 'This is the number of segments found
    Debug.Print Segments
    'Check the detected number of segments. If is =0, then the given
    'file name is not a segment file
    If Segments = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
    
    'Check if the destination file to be created does exist in the same dir
    'If yes, return error in the function return value
    DestinationFile = FPath & FNameNoExt
    If Dir(DestinationFile) <> "" Then
        ErrorCode = 3
        GoTo ErrorHandler
    End If

    'Open the destination file for binary write
    Open DestinationFile For Binary Access Write As #1 Len = 1
    
    Do
        'Increase the number of segments counter by 1
        SegmentNumber = SegmentNumber + 1
        
        'Compose the file name of the new segment file to be opened and read
        SourceFile = FPath & FNameNoExt & "." & CStr(Format(SegmentNumber, "000"))
        
        'Open the source file segment for binary read
        Open SourceFile For Binary Access Read As #2 Len = 1
        
        'Get the total number of bytes in the current segment file
        RemainingBytes = FileLen(SourceFile)
       
       'Read bytes from the source file (the current segment file) and write them to the destination file
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
       Do
            
            Select Case RemainingBytes
                Case Is >= 12000
                    'Read 12000 bytes of data from the source file
                    Get #2, , Bytes.S12000
                    'Write the bytes to the destination file
                    Put #1, , Bytes.S12000
                    'Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    'Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    'Yield to windows and other processes to do their jobs
                    'Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #2, , Bytes.S6000
                    Put #1, , Bytes.S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #2, , Bytes.S3000
                    Put #1, , Bytes.S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #2, , Bytes.S1500
                    Put #1, , Bytes.S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #2, , Bytes.S500
                    Put #1, , Bytes.S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #2, , Bytes.S100
                    Put #1, , Bytes.S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #2, , Bytes.S25
                    Put #1, , Bytes.S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #2, , Bytes.S5
                    Put #1, , Bytes.S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #2, , Bytes.S1
                    Put #1, , Bytes.S1
                    RemainingBytes = RemainingBytes - 1
                    BytesDone = BytesDone + 1
                    DoEvents
                Case Is = 0
                    'When the loop enters here, the segment bytes are completed.
                    'Close the segment file and exit the loop
                    Close 2
                    DoEvents
                    Exit Do
            End Select
            
            'Update the percent control on the form
            frmMain.spProg (Int((BytesDone / TotalBytes) * 100))
            'Refresh the form and yield to windows
            DoEvents
        Loop
        
    Loop Until SegmentNumber = Segments
    'Close the destination file
    Close 1
    
    'When the code reaches this point, everything went OK.
    'Acknowledge the number of segments, assign the value '0' to the function and exit
    NumOfSegments = Segments
    MergeFiles = 0
    frmMain.spProg (0)
    Exit Function
    
ErrorHandler:
    
    'This is entered only when an error occures
    Select Case ErrorCode
        Case Is = 0 'Unknown error
            Reset   'Close any open files
            MergeFiles = 4   'Assign error code 4 to the function
        Case Else 'Assign error code value to the function
            MergeFiles = ErrorCode
    End Select
    
    Exit Function

End Function
    

