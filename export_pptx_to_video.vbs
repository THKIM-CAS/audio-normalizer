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

                Dim waitCount
                waitCount = 0
                Dim currentStatus

                ' Wait while video is being created (status 0 or 1)
                Do
                    currentStatus = presentation.CreateVideoStatus

                    If currentStatus = 0 Or currentStatus = 1 Then
                        ' Still processing
                        waitCount = waitCount + 1
                        If waitCount Mod 5 = 0 Then
                            WScript.Echo "    Still processing... (" & waitCount & " seconds elapsed)"
                        End If
                        WScript.Sleep 1000 ' Wait 1 second
                    Else
                        ' Status changed to complete or failed
                        Exit Do
                    End If

                    ' Safety timeout: 30 minutes maximum
                    If waitCount > 1800 Then
                        WScript.Echo "    WARNING: Timeout after 30 minutes"
                        Exit Do
                    End If
                Loop

                ' Check final status
                currentStatus = presentation.CreateVideoStatus

                If currentStatus = 2 Then
                    ' Verify the file was actually created
                    If fso.FileExists(videoPath) Then
                        WScript.Echo "    SUCCESS: Video created successfully (" & waitCount & " seconds)"
                        processedFiles = processedFiles + 1
                    Else
                        WScript.Echo "    ERROR: Status shows success but video file not found"
                        failedFiles = failedFiles + 1
                    End If
                ElseIf currentStatus = 3 Then
                    WScript.Echo "    ERROR: Video creation failed (status 3)"
                    failedFiles = failedFiles + 1
                Else
                    WScript.Echo "    WARNING: Unknown final status - " & currentStatus
                    ' Check if file exists anyway
                    If fso.FileExists(videoPath) Then
                        WScript.Echo "    Note: Video file found despite unknown status"
                        processedFiles = processedFiles + 1
                    Else
                        failedFiles = failedFiles + 1
                    End If
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
