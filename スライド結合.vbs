Option Explicit

' ============================================================================
'  Slide Merger  -  merge several PowerPoint files into one
' ----------------------------------------------------------------------------
'  Usage : drag & drop one or more .pptx / .ppt / .pptm files onto this file.
'
'  What it does:
'    - Merges every slide from the dropped files into a single presentation,
'      keeping each slide's original look, by driving PowerPoint over COM.
'    - Files are merged in filename order (case-insensitive).
'    - Output goes next to the FIRST file, named:
'          <PREFIX>YYYYMMDD_HHMMSS.pptx
'    - The original files are never modified.
'
'  Requirements:
'    - Windows Script Host (built in) + Microsoft PowerPoint installed.
'    - No extra installation needed.
'
'  Note on text:
'    The script body is pure ASCII. Every piece of Japanese UI text is rebuilt
'    at run time from Unicode code points (see U()), so it can never turn into
'    mojibake no matter how this file is saved or what code page is active.
' ============================================================================

' ---- PowerPoint save format: ppSaveAsOpenXMLPresentation (.pptx) ----
Const ppSaveAsOpenXMLPresentation = 24

' ---- MsgBox flags ----
Const MB_OK              = 0
Const MB_ICONERROR       = 16
Const MB_ICONEXCLAMATION = 48
Const MB_ICONINFORMATION = 64

' ---- Japanese UI strings, rebuilt from Unicode code points ----
Dim TITLE, NOFILE, NOPPT, DONE, L_FILES, L_SLIDES, L_OUTPUT, ERRHEAD, FAILHEAD, PREFIX
TITLE    = U("30b9 30e9 30a4 30c9 7d50 5408 30c4 30fc 30eb")
NOFILE   = U("7d50 5408 3057 305f 3044 0020 0050 006f 0077 0065 0072 0050 006f 0069 006e 0074 0020 30d5 30a1 30a4 30eb 3092 3001 3053 306e 30a2 30a4 30b3 30f3 3078 307e 3068 3081 3066 30c9 30e9 30c3 30b0 ff06 30c9 30ed 30c3 30d7 3057 3066 304f 3060 3055 3044 3002")
NOPPT    = U("0050 006f 0077 0065 0072 0050 006f 0069 006e 0074 0020 3092 8d77 52d5 3067 304d 307e 305b 3093 3067 3057 305f 3002 0050 006f 0077 0065 0072 0050 006f 0069 006e 0074 0020 304c 30a4 30f3 30b9 30c8 30fc 30eb 3055 308c 3066 3044 308b 304b 78ba 8a8d 3057 3066 304f 3060 3055 3044 3002")
DONE     = U("7d50 5408 304c 5b8c 4e86 3057 307e 3057 305f 3002")
L_FILES  = U("30d5 30a1 30a4 30eb 6570 ff1a")
L_SLIDES = U("30b9 30e9 30a4 30c9 6570 ff1a")
L_OUTPUT = U("51fa 529b 5148 ff1a")
ERRHEAD  = U("30a8 30e9 30fc 304c 767a 751f 3057 307e 3057 305f ff1a")
FAILHEAD = U("6b21 306e 30d5 30a1 30a4 30eb 306f 958b 3051 307e 305b 3093 3067 3057 305f ff1a")
PREFIX   = U("7d50 5408 30b9 30e9 30a4 30c9 005f")

Main

' ----------------------------------------------------------------------------
Sub Main()
    Dim fso, files, i

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- gather the supported files that were dropped ---
    files = CollectFiles(WScript.Arguments, fso)
    If UBound(files) < 0 Then
        MsgBox NOFILE, MB_OK + MB_ICONINFORMATION, TITLE
        WScript.Quit 0
    End If

    ' --- merge in filename order ---
    SortByFileName files, fso

    ' --- decide the output path (next to the first file) ---
    Dim outFolder, outPath
    outFolder = fso.GetParentFolderName(files(0))
    outPath   = fso.BuildPath(outFolder, PREFIX & TimeStamp() & ".pptx")

    ' --- start PowerPoint ---
    Dim ppt
    On Error Resume Next
    Set ppt = CreateObject("PowerPoint.Application")
    If Err.Number <> 0 Or ppt Is Nothing Then
        On Error GoTo 0
        MsgBox NOPPT, MB_OK + MB_ICONERROR, TITLE
        WScript.Quit 1
    End If
    ppt.Visible = True
    On Error GoTo 0

    ' --- open the first file as the base, then append the rest ---
    Dim pres, failed, total
    failed = ""

    On Error Resume Next
    Set pres = ppt.Presentations.Open(files(0))
    If Err.Number <> 0 Or pres Is Nothing Then
        Dim openErr : openErr = Err.Description
        On Error GoTo 0
        MsgBox ERRHEAD & vbCrLf & openErr, MB_OK + MB_ICONERROR, TITLE
        WScript.Quit 1
    End If
    On Error GoTo 0

    ' append every slide from the remaining files, keeping their look
    For i = 1 To UBound(files)
        On Error Resume Next
        Err.Clear
        pres.Slides.InsertFromFile files(i), pres.Slides.Count
        If Err.Number <> 0 Then
            failed = failed & "  - " & fso.GetFileName(files(i)) & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    Next

    ' --- save the merged result (originals are untouched) ---
    On Error Resume Next
    Err.Clear
    pres.SaveAs outPath, ppSaveAsOpenXMLPresentation
    If Err.Number <> 0 Then
        Dim saveErr : saveErr = Err.Description
        On Error GoTo 0
        MsgBox ERRHEAD & vbCrLf & saveErr, MB_OK + MB_ICONERROR, TITLE
        WScript.Quit 1
    End If
    On Error GoTo 0

    total = pres.Slides.Count

    ' --- report success; leave the merged file open for review ---
    Dim msg
    msg = DONE & vbCrLf & vbCrLf
    msg = msg & L_FILES  & (UBound(files) + 1) & vbCrLf
    msg = msg & L_SLIDES & total & vbCrLf
    msg = msg & L_OUTPUT & vbCrLf & outPath
    If Len(failed) > 0 Then
        msg = msg & vbCrLf & vbCrLf & FAILHEAD & vbCrLf & failed
    End If
    MsgBox msg, MB_OK + MB_ICONINFORMATION, TITLE
End Sub

' ----------------------------------------------------------------------------
' Collect dropped paths that are real files with a PowerPoint extension.
Function CollectFiles(args, fso)
    Dim arr(), n, i, p, ext
    ReDim arr(-1)
    n = 0
    For i = 0 To args.Count - 1
        p = args(i)
        If fso.FileExists(p) Then
            ext = LCase(fso.GetExtensionName(p))
            If ext = "pptx" Or ext = "ppt" Or ext = "pptm" Then
                ReDim Preserve arr(n)
                arr(n) = p
                n = n + 1
            End If
        End If
    Next
    CollectFiles = arr
End Function

' ----------------------------------------------------------------------------
' Insertion sort by leaf filename, case-insensitive. Modifies the array in place.
Sub SortByFileName(arr, fso)
    Dim i, j, keyVal, keyName
    For i = 1 To UBound(arr)
        keyVal  = arr(i)
        keyName = fso.GetFileName(keyVal)
        j = i - 1
        Do While j >= 0
            If StrComp(fso.GetFileName(arr(j)), keyName, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = keyVal
    Next
End Sub

' ----------------------------------------------------------------------------
' "YYYYMMDD_HHMMSS" from the current local time.
Function TimeStamp()
    Dim n
    n = Now
    TimeStamp = Year(n) & P2(Month(n)) & P2(Day(n)) & "_" & _
                P2(Hour(n)) & P2(Minute(n)) & P2(Second(n))
End Function

Function P2(v)
    P2 = Right("0" & v, 2)
End Function

' ----------------------------------------------------------------------------
' Build a Unicode string from space-separated hex code points (keeps the
' source file pure ASCII so the Japanese UI never becomes mojibake).
Function U(s)
    Dim parts, i, r
    parts = Split(s, " ")
    r = ""
    For i = 0 To UBound(parts)
        r = r & ChrW(CLng("&H" & parts(i)))
    Next
    U = r
End Function
