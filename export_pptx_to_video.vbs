' VBScript to batch export PowerPoint presentations to video files
' Usage: cscript export_pptx_to_video.vbs [directory_path] [output_directory]
' If no arguments provided, it will prompt for directory paths

Option Explicit

Dim fso, objPPT, folder, file, args
Dim inputFolder, outputFolder
Dim pptxFiles, presentation
Dim videoPath, baseName
Dim totalFiles, processedFiles, failedFiles
Dim videoQuality ' Quality: 1=Low, 2=Medium, 3=High, 4=Very High, 5=HD 720p, 6=Full HD 1080p

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments
totalFiles = 0
processedFiles = 0
failedFiles = 0

' Video quality setting (can be changed: 1-6, default is 5 for HD 720p)
videoQuality = 5

' Get input folder path
If args.Count >= 1 Then
    inputFolder = args(0)
Else
    inputFolder = InputBox("Enter the directory path containing PowerPoint files:", "Input Directory")
    If inputFolder = "" Then
        WScript.Echo "No input directory specified. Exiting."
        WScript.Quit
    End If
End If

' Validate input folder
If Not fso.FolderExists(inputFolder) Then
    WScript.Echo "Error: Input folder does not exist: " & inputFolder
    WScript.Quit
End If

' Get output folder path
If args.Count >= 2 Then
    outputFolder = args(1)
Else
    outputFolder = InputBox("Enter the output directory for video files (leave empty to use same as input):", "Output Directory", inputFolder)
    If outputFolder = "" Then
        outputFolder = inputFolder
    End If
End If

' Create output folder if it doesn't exist
If Not fso.FolderExists(outputFolder) Then
    fso.CreateFolder(outputFolder)
End If

WScript.Echo "========================================="
WScript.Echo "PowerPoint to Video Batch Exporter"
WScript.Echo "========================================="
WScript.Echo "Input folder:  " & inputFolder
WScript.Echo "Output folder: " & outputFolder
WScript.Echo "Video quality: " & GetQualityName(videoQuality)
WScript.Echo "========================================="
WScript.Echo ""

' Create PowerPoint application object
On Error Resume Next
Set objPPT = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not create PowerPoint application object."
    WScript.Echo "Make sure Microsoft PowerPoint is installed."
    WScript.Quit
End If
On Error Goto 0

' Note: PowerPoint does not allow Visible = False via COM automation
' The application window must remain visible, but we minimize it
objPPT.Visible = True
objPPT.WindowState = 2 ' ppWindowMinimized = 2

' Get all .pptx and .ppt files in the folder
Set folder = fso.GetFolder(inputFolder)

' Count total files first
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "pptx" Or _
       LCase(fso.GetExtensionName(file.Name)) = "ppt" Then
        totalFiles = totalFiles + 1
    End If
Next

If totalFiles = 0 Then
    WScript.Echo "No PowerPoint files (.pptx or .ppt) found in the specified directory."
    objPPT.Quit
    Set objPPT = Nothing
    WScript.Quit
End If

WScript.Echo "Found " & totalFiles & " PowerPoint file(s) to process."
WScript.Echo ""

' Process each PowerPoint file
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "pptx" Or _
       LCase(fso.GetExtensionName(file.Name)) = "ppt" Then

        baseName = fso.GetBaseName(file.Name)
        videoPath = fso.BuildPath(outputFolder, baseName & ".mp4")

        WScript.Echo "[" & (processedFiles + failedFiles + 1) & "/" & totalFiles & "] Processing: " & file.Name

        On Error Resume Next

        ' Open the presentation
        Set presentation = objPPT.Presentations.Open(file.Path, , , False)

        If Err.Number <> 0 Then
            WScript.Echo "    ERROR: Could not open file - " & Err.Description
            failedFiles = failedFiles + 1
            Err.Clear
        Else
            ' Export to video
            ' CreateVideo parameters:
            ' 1. FileName - output path
            ' 2. UseTimingsAndNarrations - True to use slide timings, False to use default
            ' 3. DefaultSlideDuration - seconds per slide if not using timings (default 5)
            ' 4. VertResolution - vertical resolution in pixels (720, 1080, etc.)
            ' 5. FramesPerSecond - frames per second (30 or 60)
            ' 6. Quality - video quality (1-100, default 85)

            Dim vertRes, fps, useTimings, slideDuration

            ' Set parameters based on quality level
            Select Case videoQuality
                Case 1 ' Low quality
                    vertRes = 480
                    fps = 24
                Case 2 ' Medium quality
                    vertRes = 576
                    fps = 25
                Case 3 ' High quality
                    vertRes = 720
                    fps = 25
                Case 4 ' Very High quality
                    vertRes = 720
                    fps = 30
                Case 5 ' HD 720p
                    vertRes = 720
                    fps = 30
                Case 6 ' Full HD 1080p
                    vertRes = 1080
                    fps = 30
                Case Else ' Default to HD 720p
                    vertRes = 720
                    fps = 30
            End Select

            useTimings = True
            slideDuration = 5 ' Default 5 seconds per slide if no timings

            ' Check if presentation has slide timings
            Dim hasTimings
            hasTimings = False
            Dim sld
            For Each sld In presentation.Slides
                If sld.SlideShowTransition.AdvanceOnTime Then
                    hasTimings = True
                    Exit For
                End If
            Next

            If Not hasTimings Then
                useTimings = False
                WScript.Echo "    Note: No slide timings found, using " & slideDuration & " seconds per slide"
            End If

            WScript.Echo "    Exporting to video: " & baseName & ".mp4"
            WScript.Echo "    Resolution: " & vertRes & "p @ " & fps & " fps"

            ' Create the video
            presentation.CreateVideo videoPath, useTimings, slideDuration, vertRes, fps, 85

            If Err.Number <> 0 Then
                WScript.Echo "    ERROR: Could not create video - " & Err.Description
                failedFiles = failedFiles + 1
                Err.Clear
            Else
                ' Wait for video creation to complete
                WScript.Echo "    Please wait... Video is being created..."
                WScript.Echo "    (This may take several minutes depending on presentation size)"

                ' Note: CreateVideo is asynchronous, we need to wait for it to complete
                ' CreateVideoStatus: 0=NotStarted, 1=InProgress, 2=Complete, 3=Failed

                ' Give it a moment to start
                WScript.Sleep 2000

                Dim waitCount, checkCount
                Dim lastFileSize, currentFileSize
                Dim stableCount
                waitCount = 0
                checkCount = 0
                lastFileSize = 0
                stableCount = 0
                Dim currentStatus

                ' Wait while video is being created
                ' Strategy: Monitor both status AND file size changes
                Do
                    currentStatus = presentation.CreateVideoStatus

                    ' Check if video file exists and its size
                    If fso.FileExists(videoPath) Then
                        currentFileSize = fso.GetFile(videoPath).Size

                        ' If file size hasn't changed in 10 checks (10 seconds), consider it complete
                        If currentFileSize = lastFileSize And currentFileSize > 0 Then
                            stableCount = stableCount + 1
                            If stableCount >= 10 Then
                                WScript.Echo "    Video file size stable, export appears complete"
                                Exit Do
                            End If
                        Else
                            stableCount = 0 ' Reset if size changed
                        End If

                        lastFileSize = currentFileSize
                    End If

                    ' Also check status
                    If currentStatus = 2 Then
                        ' Complete status
                        WScript.Echo "    Status indicates completion"
                        Exit Do
                    ElseIf currentStatus = 3 Then
                        ' Failed status
                        WScript.Echo "    Status indicates failure"
                        Exit Do
                    End If

                    ' Progress indicator
                    waitCount = waitCount + 1
                    If waitCount Mod 10 = 0 Then
                        If fso.FileExists(videoPath) Then
                            WScript.Echo "    Still processing... (" & waitCount & " seconds, file size: " & FormatBytes(currentFileSize) & ")"
                        Else
                            WScript.Echo "    Still processing... (" & waitCount & " seconds)"
                        End If
                    End If

                    WScript.Sleep 1000 ' Wait 1 second

                    ' Safety timeout: 2 hours maximum (some large presentations take a long time)
                    If waitCount > 7200 Then
                        WScript.Echo "    WARNING: Timeout after 2 hours"
                        Exit Do
                    End If
                Loop

                ' Check final result
                currentStatus = presentation.CreateVideoStatus

                If fso.FileExists(videoPath) Then
                    Dim finalFileSize
                    finalFileSize = fso.GetFile(videoPath).Size

                    If finalFileSize > 0 Then
                        WScript.Echo "    SUCCESS: Video created successfully (" & waitCount & " seconds, " & FormatBytes(finalFileSize) & ")"
                        processedFiles = processedFiles + 1
                    Else
                        WScript.Echo "    ERROR: Video file exists but is empty"
                        failedFiles = failedFiles + 1
                    End If
                Else
                    WScript.Echo "    ERROR: Video file not found (Status: " & currentStatus & ")"
                    failedFiles = failedFiles + 1
                End If
            End If

            ' Close the presentation without saving
            presentation.Close
            Set presentation = Nothing
        End If

        On Error Goto 0
        WScript.Echo ""
    End If
Next

' Summary
WScript.Echo "========================================="
WScript.Echo "Batch Export Complete"
WScript.Echo "========================================="
WScript.Echo "Total files found:     " & totalFiles
WScript.Echo "Successfully exported: " & processedFiles
WScript.Echo "Failed:                " & failedFiles
WScript.Echo "========================================="

' Clean up
objPPT.Quit
Set objPPT = Nothing
Set fso = Nothing

WScript.Echo ""
WScript.Echo "Press Enter to exit..."
WScript.StdIn.ReadLine

' Function to get quality name
Function GetQualityName(quality)
    Select Case quality
        Case 1
            GetQualityName = "Low (480p, 24fps)"
        Case 2
            GetQualityName = "Medium (576p, 25fps)"
        Case 3
            GetQualityName = "High (720p, 25fps)"
        Case 4
            GetQualityName = "Very High (720p, 30fps)"
        Case 5
            GetQualityName = "HD 720p (30fps)"
        Case 6
            GetQualityName = "Full HD 1080p (30fps)"
        Case Else
            GetQualityName = "HD 720p (30fps)"
    End Select
End Function

' Function to format bytes into human-readable format
Function FormatBytes(bytes)
    Dim size
    size = CDbl(bytes)

    If size < 1024 Then
        FormatBytes = size & " B"
    ElseIf size < 1048576 Then
        FormatBytes = Round(size / 1024, 2) & " KB"
    ElseIf size < 1073741824 Then
        FormatBytes = Round(size / 1048576, 2) & " MB"
    Else
        FormatBytes = Round(size / 1073741824, 2) & " GB"
    End If
End Function
