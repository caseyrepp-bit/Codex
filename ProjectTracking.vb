' ============================================
' SINGLE PROJECT EXPORT TOOL - COMPLETE FIXED
' Pulls one project and all its data
' Exports to Excel and PowerPoint
' ============================================

Option Explicit

' Color Constants
Const PRIMARY_RED As Long = 9109504       ' RGB(128, 0, 0) Deep Red
Const SECONDARY_GREY As Long = 6710886    ' RGB(102, 102, 102) Grey
Const LIGHT_GREY As Long = 14277081       ' RGB(217, 217, 217) Light Grey
Const STATUS_GREEN As Long = 13561798     ' RGB(198, 239, 206)
Const STATUS_YELLOW As Long = 10284031    ' RGB(255, 235, 156)
Const STATUS_RED As Long = 13486335       ' RGB(255, 199, 206)

' Status ID mappings (from your Freshservice instance)
Const STATUS_OPEN_ID As String = "5000059610"
Const STATUS_IN_PROGRESS_ID As String = "5000059611"
Const STATUS_CLOSED_ID As String = "5000059612"

' Type ID mappings (from your Freshservice instance)
Const TYPE_FOLDER_ID As String = "5000119020"
Const TYPE_PROJECT_ID As String = "5000119023"
Const TYPE_TASK_ID As String = "5000119024"

' ============================================
' CREATE BUTTONS - Run once to add buttons
' ============================================
Sub CreateButtons()
    
    Dim ws As Worksheet
    Dim btn As Button
    
    Set ws = ThisWorkbook.Sheets("Config")
    
    ' Clear existing buttons
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp
    
    ' Select & Load Project Button
    Set btn = ws.Buttons.Add(200, 10, 140, 36)
    With btn
        .Name = "btnSelectProject"
        .Caption = "Select Project"
        .OnAction = "SelectAndLoadProject"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' Export to PowerPoint Button
    Set btn = ws.Buttons.Add(350, 10, 150, 36)
    With btn
        .Name = "btnExport"
        .Caption = "Export to PowerPoint"
        .OnAction = "ExportProjectToPowerPoint"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    MsgBox "Buttons created!", vbInformation
    
End Sub

' ============================================
' SELECT AND LOAD A SINGLE PROJECT
' ============================================
Sub SelectAndLoadProject()
    
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    
    Dim apiKey As String
    Dim domain As String
    
    apiKey = Trim(CStr(wsConfig.Range("B1").Value))
    domain = Trim(CStr(wsConfig.Range("B2").Value))
    domain = Replace(domain, "https://", "")
    domain = Replace(domain, "http://", "")
    If Right(domain, 1) = "/" Then domain = Left(domain, Len(domain) - 1)
    
    If apiKey = "" Or domain = "" Then
        MsgBox "Please enter API Key in B1 and Domain in B2", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "Fetching project list..."
    
    ' Get list of projects
    Dim projUrl As String
    projUrl = "https://" & domain & "/api/v2/pm/projects"
    
    Dim projJson As String
    projJson = MakeAPICall(projUrl, apiKey)
    
    If Left(projJson, 5) = "ERROR" Then
        MsgBox "Failed to fetch projects: " & projJson, vbCritical
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' Build project list
    Dim projectList As String
    Dim projIDs As New Collection
    Dim projNames As New Collection
    Dim searchPos As Long
    Dim idx As Long
    
    searchPos = 1
    idx = 0
    
    Do
        Dim idPos As Long
        idPos = InStr(searchPos, projJson, """id"":")
        If idPos = 0 Then Exit Do
        
        Dim idVal As String
        idVal = ExtractNumber(projJson, idPos + 5)
        
        Dim nameVal As String
        nameVal = ExtractString(projJson, idPos, "name")
        
        If idVal <> "" And nameVal <> "" Then
            idx = idx + 1
            projIDs.Add idVal
            projNames.Add nameVal
            projectList = projectList & idx & ". " & nameVal & vbCrLf
        End If
        
        searchPos = idPos + 20
        If idx > 50 Then Exit Do
    Loop
    
    Application.StatusBar = False
    
    If idx = 0 Then
        MsgBox "No projects found.", vbExclamation
        Exit Sub
    End If
    
    ' Ask user to select
    Dim selection As String
    selection = InputBox("Select a project number:" & vbCrLf & vbCrLf & projectList, "Select Project", "1")
    
    If selection = "" Then Exit Sub
    
    Dim selectedIdx As Long
    On Error Resume Next
    selectedIdx = CLng(selection)
    On Error GoTo 0
    
    If selectedIdx < 1 Or selectedIdx > projIDs.count Then
        MsgBox "Invalid selection.", vbExclamation
        Exit Sub
    End If
    
    ' Load the selected project
    LoadSingleProject projIDs(selectedIdx), projNames(selectedIdx), apiKey, domain
    
End Sub

' ============================================
' LOAD SINGLE PROJECT DATA
' ============================================
Sub LoadSingleProject(projectID As String, projectName As String, apiKey As String, domain As String)
    
    Dim wsProject As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSummary As Worksheet
    
    Set wsProject = GetOrCreateSheet("Project")
    Set wsTasks = GetOrCreateSheet("Tasks")
    Set wsSummary = GetOrCreateSheet("Summary")
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Loading project: " & projectName
    
    ' Clear sheets
    wsProject.Cells.Clear
    wsTasks.Cells.Clear
    wsSummary.Cells.Clear
    
    ' ============================================
    ' FETCH PROJECT DETAILS
    ' ============================================
    
    Dim detailUrl As String
    detailUrl = "https://" & domain & "/api/v2/pm/projects/" & projectID
    
    Dim detailJson As String
    detailJson = MakeAPICall(detailUrl, apiKey)
    
    ' Setup Project sheet
    wsProject.Range("A1").Value = "PROJECT DETAILS"
    wsProject.Range("A1").Font.Bold = True
    wsProject.Range("A1").Font.Size = 14
    wsProject.Range("A1").Font.Color = PRIMARY_RED
    
    wsProject.Range("A3").Value = "Project ID"
    wsProject.Range("B3").Value = projectID
    wsProject.Range("A4").Value = "Project Name"
    wsProject.Range("B4").Value = projectName
    
    Dim row As Long
    row = 5
    
    If Left(detailJson, 5) <> "ERROR" Then
        wsProject.Range("A" & row).Value = "Key"
        wsProject.Range("B" & row).Value = ExtractString(detailJson, 1, "key")
        row = row + 1
        
        wsProject.Range("A" & row).Value = "Status"
        Dim projStatusId As String
        projStatusId = ExtractNumber(detailJson, InStr(detailJson, """status_id"":") + 12)
        wsProject.Range("B" & row).Value = GetStatusName(projStatusId)
        wsProject.Range("B" & row).Interior.Color = GetStatusColor(projStatusId)
        row = row + 1
        
        wsProject.Range("A" & row).Value = "Priority"
        Dim priorityId As String
        priorityId = ExtractNumber(detailJson, InStr(detailJson, """priority_id"":") + 14)
        wsProject.Range("B" & row).Value = GetPriorityName(priorityId)
        row = row + 1
        
        wsProject.Range("A" & row).Value = "Start Date"
        wsProject.Range("B" & row).Value = FormatAPIDate(ExtractString(detailJson, 1, "start_date"))
        row = row + 1
        
        wsProject.Range("A" & row).Value = "End Date"
        wsProject.Range("B" & row).Value = FormatAPIDate(ExtractString(detailJson, 1, "end_date"))
        row = row + 1
        
        wsProject.Range("A" & row).Value = "Created"
        wsProject.Range("B" & row).Value = FormatAPIDate(ExtractString(detailJson, 1, "created_at"))
        row = row + 1
        
        wsProject.Range("A" & row).Value = "Updated"
        wsProject.Range("B" & row).Value = FormatAPIDate(ExtractString(detailJson, 1, "updated_at"))
        row = row + 1
        
        wsProject.Range("A" & row).Value = "Owner ID"
        wsProject.Range("B" & row).Value = ExtractNumber(detailJson, InStr(detailJson, """owner_id"":") + 11)
        row = row + 1
    End If
    
    ' Format Project sheet
    wsProject.Range("A3:A" & row - 1).Font.Bold = True
    wsProject.Range("A3:A" & row - 1).Interior.Color = PRIMARY_RED
    wsProject.Range("A3:A" & row - 1).Font.Color = RGB(255, 255, 255)
    wsProject.Columns("A:B").AutoFit
    
    ' ============================================
    ' FETCH ALL TASKS
    ' ============================================
    
    Application.StatusBar = "Loading tasks..."
    
    Dim taskUrl As String
    taskUrl = "https://" & domain & "/api/v2/pm/projects/" & projectID & "/tasks"
    
    Dim taskJson As String
    taskJson = MakeAPICall(taskUrl, apiKey)
    
    ' ============================================
    ' PARSE TASKS INTO ARRAYS FOR SORTING
    ' ============================================
    
    Dim taskData() As Variant
    Dim taskCount As Long
    taskCount = 0
    
    ' First pass: count tasks
    Dim countPos As Long
    countPos = 1
    Do
        Dim countIdPos As Long
        countIdPos = InStr(countPos, taskJson, """id"":")
        If countIdPos = 0 Then Exit Do
        taskCount = taskCount + 1
        countPos = countIdPos + 10
        If taskCount > 500 Then Exit Do
    Loop
    
    If taskCount = 0 Then
        ' No tasks - still show summary
        BuildSummarySheet wsSummary, projectName, projectID, 0, 0, 0, 0
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "Project loaded but no tasks found.", vbInformation
        wsSummary.Activate
        Exit Sub
    End If
    
    ' Dimension array: ID, Title, Type, TypeID, Status, StatusID, OwnerID, StartDate, DueDate, Duration, ParentID, Created, Updated, Level
    ReDim taskData(1 To taskCount, 1 To 14)
    
    Dim openCount As Long, inProgressCount As Long, closedCount As Long
    Dim taskSearchPos As Long
    Dim idx As Long
    taskSearchPos = 1
    idx = 0
    
    ' Second pass: extract all task data
    Do
        Dim taskIdPos As Long
        taskIdPos = InStr(taskSearchPos, taskJson, """id"":")
        If taskIdPos = 0 Then Exit Do
        
        Dim taskId As String
        taskId = ExtractNumber(taskJson, taskIdPos + 5)
        
        Dim taskTitle As String
        taskTitle = ExtractString(taskJson, taskIdPos, "title")
        
        If taskId <> "" And taskTitle <> "" Then
            idx = idx + 1
            
            ' Extract status_id (FIXED)
            Dim taskStatusId As String
            Dim statusPos As Long
            statusPos = InStr(taskIdPos, taskJson, """status_id"":")
            If statusPos > 0 And statusPos < taskIdPos + 500 Then
                taskStatusId = ExtractNumber(taskJson, statusPos + 12)
            Else
                taskStatusId = ""
            End If
            
            ' Extract type_id (FIXED)
            Dim taskTypeId As String
            Dim typePos As Long
            typePos = InStr(taskIdPos, taskJson, """type_id"":")
            If typePos > 0 And typePos < taskIdPos + 500 Then
                taskTypeId = ExtractNumber(taskJson, typePos + 10)
            Else
                taskTypeId = ""
            End If
            
            ' Extract parent_id
            Dim parentId As String
            Dim parentPos As Long
            parentPos = InStr(taskIdPos, taskJson, """parent_id"":")
            If parentPos > 0 And parentPos < taskIdPos + 800 Then
                parentId = ExtractNumber(taskJson, parentPos + 12)
                If parentId = "null" Or parentId = "" Then parentId = ""
            Else
                parentId = ""
            End If
            
            ' Extract planned_start_date (FIXED)
            Dim startDate As String
            startDate = ExtractString(taskJson, taskIdPos, "planned_start_date")
            
            ' Extract due date (FIXED)
            Dim dueDate As String
            dueDate = ExtractString(taskJson, taskIdPos, "planned_end_date")
            
            ' Extract assignee
            Dim assigneeId As String
            Dim assigneePos As Long
            assigneePos = InStr(taskIdPos, taskJson, """assignee_id"":")
            If assigneePos > 0 And assigneePos < taskIdPos + 500 Then
                assigneeId = ExtractNumber(taskJson, assigneePos + 14)
            Else
                assigneeId = ""
            End If
            
            ' Store in array
            taskData(idx, 1) = taskId
            taskData(idx, 2) = taskTitle
            taskData(idx, 3) = GetTypeName(taskTypeId)
            taskData(idx, 4) = taskTypeId
            taskData(idx, 5) = GetTaskStatusName(taskStatusId)
            taskData(idx, 6) = taskStatusId
            taskData(idx, 7) = assigneeId
            taskData(idx, 8) = FormatAPIDate(startDate)
            taskData(idx, 9) = FormatAPIDate(dueDate)
            taskData(idx, 10) = ExtractString(taskJson, taskIdPos, "planned_duration")
            taskData(idx, 11) = parentId
            taskData(idx, 12) = FormatAPIDate(ExtractString(taskJson, taskIdPos, "created_at"))
            taskData(idx, 13) = FormatAPIDate(ExtractString(taskJson, taskIdPos, "updated_at"))
            taskData(idx, 14) = 0 ' Level - will be calculated
            
            ' Count by status
            Select Case taskStatusId
                Case STATUS_OPEN_ID: openCount = openCount + 1
                Case STATUS_IN_PROGRESS_ID: inProgressCount = inProgressCount + 1
                Case STATUS_CLOSED_ID: closedCount = closedCount + 1
            End Select
        End If
        
        taskSearchPos = taskIdPos + 50
        If idx >= taskCount Then Exit Do
    Loop
    
    taskCount = idx ' Actual count
    
    ' ============================================
    ' BUILD HIERARCHY AND SORT
    ' ============================================
    
    Application.StatusBar = "Building hierarchy..."
    
    ' Calculate levels for each task
    Dim i As Long
    For i = 1 To taskCount
        taskData(i, 14) = GetTaskLevel(CStr(taskData(i, 1)), CStr(taskData(i, 11)), taskData, taskCount)
    Next i
    
    ' Build sorted output with hierarchy
    Dim sortedTasks() As Long
    ReDim sortedTasks(1 To taskCount)
    Dim sortedCount As Long
    sortedCount = 0
    
    ' First, find all root items (no parent) and add them with their children
    For i = 1 To taskCount
        If CStr(taskData(i, 11)) = "" Or CStr(taskData(i, 11)) = "0" Then
            AddTaskAndChildren CStr(taskData(i, 1)), taskData, taskCount, sortedTasks, sortedCount
        End If
    Next i
    
    ' ============================================
    ' WRITE TO TASKS SHEET
    ' ============================================
    
    ' Setup Tasks sheet headers
    wsTasks.Range("A1").Value = "TASKS FOR: " & projectName
    wsTasks.Range("A1").Font.Bold = True
    wsTasks.Range("A1").Font.Size = 14
    wsTasks.Range("A1").Font.Color = PRIMARY_RED
    
    wsTasks.Range("A3:L3").Value = Array("Task ID", "Title", "Type", "Status", "Owner ID", "Start Date", "Due Date", "Duration", "Parent ID", "Created", "Updated", "Level")
    FormatHeader wsTasks.Range("A3:L3")
    
    Dim taskRow As Long
    taskRow = 4
    
    For i = 1 To sortedCount
        Dim srcIdx As Long
        srcIdx = sortedTasks(i)
        
        wsTasks.Cells(taskRow, 1).Value = taskData(srcIdx, 1)  ' ID
        
        ' Add indentation based on level
        Dim indent As String
        Dim level As Long
        level = CLng(taskData(srcIdx, 14))
        indent = String(level * 3, " ")
        If level > 0 Then indent = indent & Chr(187) & " "  ' >> symbol
        
        wsTasks.Cells(taskRow, 2).Value = indent & taskData(srcIdx, 2)  ' Title with indent
        If level > 0 Then wsTasks.Cells(taskRow, 2).Font.Color = SECONDARY_GREY
        If level > 1 Then wsTasks.Cells(taskRow, 2).Font.Italic = True
        
        wsTasks.Cells(taskRow, 3).Value = taskData(srcIdx, 3)  ' Type name
        wsTasks.Cells(taskRow, 4).Value = taskData(srcIdx, 5)  ' Status name
        wsTasks.Cells(taskRow, 4).Interior.Color = GetTaskStatusColor(CStr(taskData(srcIdx, 6)))
        wsTasks.Cells(taskRow, 5).Value = taskData(srcIdx, 7)  ' Owner ID
        wsTasks.Cells(taskRow, 6).Value = taskData(srcIdx, 8)  ' Start Date
        wsTasks.Cells(taskRow, 7).Value = taskData(srcIdx, 9)  ' Due Date
        wsTasks.Cells(taskRow, 8).Value = taskData(srcIdx, 10) ' Duration
        wsTasks.Cells(taskRow, 9).Value = taskData(srcIdx, 11) ' Parent ID
        wsTasks.Cells(taskRow, 10).Value = taskData(srcIdx, 12) ' Created
        wsTasks.Cells(taskRow, 11).Value = taskData(srcIdx, 13) ' Updated
        wsTasks.Cells(taskRow, 12).Value = level ' Level
        
        taskRow = taskRow + 1
    Next i
    
    wsTasks.Columns("A:L").AutoFit
    If wsTasks.Columns("B").ColumnWidth > 60 Then wsTasks.Columns("B").ColumnWidth = 60
    
    ' ============================================
    ' BUILD SUMMARY SHEET
    ' ============================================
    
    BuildSummarySheet wsSummary, projectName, projectID, taskCount, openCount, inProgressCount, closedCount
    
    ' Update Config sheet
    ThisWorkbook.Sheets("Config").Range("B3").Value = Now
    ThisWorkbook.Sheets("Config").Range("B4").Value = projectName
    ThisWorkbook.Sheets("Config").Range("A4").Value = "Loaded Project:"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Project loaded successfully!" & vbCrLf & vbCrLf & _
           "Project: " & projectName & vbCrLf & _
           "Tasks: " & taskCount & vbCrLf & _
           "Open: " & openCount & vbCrLf & _
           "In Progress: " & inProgressCount & vbCrLf & _
           "Closed: " & closedCount & vbCrLf & _
           "Completion: " & Format((closedCount / IIf(taskCount = 0, 1, taskCount)) * 100, "0.0") & "%", vbInformation, "Complete"
    
    wsSummary.Activate
    
End Sub

' ============================================
' BUILD SUMMARY SHEET
' ============================================
Sub BuildSummarySheet(wsSummary As Worksheet, projectName As String, projectID As String, taskCount As Long, openCount As Long, inProgressCount As Long, closedCount As Long)
    
    Application.StatusBar = "Building summary..."
    
    Dim completionRate As Double
    If taskCount > 0 Then completionRate = (closedCount / taskCount) * 100
    
    wsSummary.Range("A1").Value = "PROJECT SUMMARY"
    wsSummary.Range("A1").Font.Bold = True
    wsSummary.Range("A1").Font.Size = 16
    wsSummary.Range("A1").Font.Color = PRIMARY_RED
    
    wsSummary.Range("A3").Value = "Project:"
    wsSummary.Range("B3").Value = projectName
    wsSummary.Range("B3").Font.Bold = True
    wsSummary.Range("B3").Font.Size = 14
    
    wsSummary.Range("A4").Value = "Last Refresh:"
    wsSummary.Range("B4").Value = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    
    wsSummary.Range("A5").Value = "Project ID:"
    wsSummary.Range("B5").Value = projectID
    
    wsSummary.Range("A7").Value = "COMPLETION"
    wsSummary.Range("A7").Font.Bold = True
    wsSummary.Range("A7:B7").Interior.Color = PRIMARY_RED
    wsSummary.Range("A7:B7").Font.Color = RGB(255, 255, 255)
    
    wsSummary.Range("A8").Value = "Completion Rate:"
    wsSummary.Range("B8").Value = Format(completionRate, "0.0") & "%"
    wsSummary.Range("B8").Font.Bold = True
    wsSummary.Range("B8").Font.Size = 18
    wsSummary.Range("B8").Font.Color = PRIMARY_RED
    
    wsSummary.Range("A10").Value = "TASK BREAKDOWN"
    wsSummary.Range("A10").Font.Bold = True
    wsSummary.Range("A10:B10").Interior.Color = PRIMARY_RED
    wsSummary.Range("A10:B10").Font.Color = RGB(255, 255, 255)
    
    wsSummary.Range("A11").Value = "Total Tasks:"
    wsSummary.Range("B11").Value = taskCount
    
    wsSummary.Range("A12").Value = "Open:"
    wsSummary.Range("B12").Value = openCount
    wsSummary.Range("B12").Interior.Color = LIGHT_GREY
    
    wsSummary.Range("A13").Value = "In Progress:"
    wsSummary.Range("B13").Value = inProgressCount
    wsSummary.Range("B13").Interior.Color = STATUS_YELLOW
    
    wsSummary.Range("A14").Value = "Closed:"
    wsSummary.Range("B14").Value = closedCount
    wsSummary.Range("B14").Interior.Color = STATUS_GREEN
    
    wsSummary.Range("A3:A14").Font.Bold = True
    wsSummary.Columns("A:B").AutoFit
    
End Sub

' ============================================
' HIERARCHY HELPER FUNCTIONS
' ============================================
Function GetTaskLevel(taskId As String, parentId As String, taskData() As Variant, taskCount As Long) As Long
    If parentId = "" Or parentId = "0" Then
        GetTaskLevel = 0
        Exit Function
    End If
    
    Dim level As Long
    Dim currentParent As String
    currentParent = parentId
    level = 0
    
    Do While currentParent <> "" And currentParent <> "0" And level < 10
        level = level + 1
        Dim i As Long
        Dim found As Boolean
        found = False
        For i = 1 To taskCount
            If CStr(taskData(i, 1)) = currentParent Then
                currentParent = CStr(taskData(i, 11))
                found = True
                Exit For
            End If
        Next i
        If Not found Then Exit Do
    Loop
    
    GetTaskLevel = level
End Function

Sub AddTaskAndChildren(taskId As String, taskData() As Variant, taskCount As Long, sortedTasks() As Long, ByRef sortedCount As Long)
    Dim i As Long
    For i = 1 To taskCount
        If CStr(taskData(i, 1)) = taskId Then
            sortedCount = sortedCount + 1
            sortedTasks(sortedCount) = i
            Exit For
        End If
    Next i
    
    For i = 1 To taskCount
        If CStr(taskData(i, 11)) = taskId Then
            AddTaskAndChildren CStr(taskData(i, 1)), taskData, taskCount, sortedTasks, sortedCount
        End If
    Next i
End Sub

' ============================================
' EXPORT TO POWERPOINT (FULL ORIGINAL)
' ============================================
Sub ExportProjectToPowerPoint()
    
    Dim wsProject As Worksheet
    Dim wsTasks As Worksheet
    Dim wsSummary As Worksheet
    
    On Error Resume Next
    Set wsProject = ThisWorkbook.Sheets("Project")
    Set wsTasks = ThisWorkbook.Sheets("Tasks")
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0
    
    If wsProject Is Nothing Or wsTasks Is Nothing Or wsSummary Is Nothing Then
        MsgBox "Please select and load a project first.", vbExclamation
        Exit Sub
    End If
    
    Dim projectName As String
    projectName = wsSummary.Range("B3").Value
    
    If projectName = "" Then
        MsgBox "No project data found. Please select a project first.", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "Creating PowerPoint..."
    
    ' Get task counts from summary
    Dim taskCount As Long, openCount As Long, inProgressCount As Long, closedCount As Long
    taskCount = CLng(wsSummary.Range("B11").Value)
    openCount = CLng(wsSummary.Range("B12").Value)
    inProgressCount = CLng(wsSummary.Range("B13").Value)
    closedCount = CLng(wsSummary.Range("B14").Value)
    
    Dim completionRate As Double
    If taskCount > 0 Then completionRate = (closedCount / taskCount) * 100
    
    ' Create PowerPoint
    Dim pptApp As Object
    Dim ppt As Object
    Dim sld As Object
    Dim shp As Object
    Dim tbl As Object
    
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    If pptApp Is Nothing Then
        MsgBox "Could not start PowerPoint.", vbCritical
        Exit Sub
    End If
    
    pptApp.Visible = True
    Set ppt = pptApp.Presentations.Add
    
    Dim slideIndex As Long
    Dim i As Long, j As Long, r As Long
    slideIndex = 0
    
    ' ----- SLIDE 1: TITLE -----
    slideIndex = slideIndex + 1
    Set sld = ppt.Slides.Add(slideIndex, 1)
    
    With sld.Shapes(1).TextFrame.TextRange
        .text = projectName
        .Font.Name = "Arial"
        .Font.Size = 40
        .Font.Bold = True
        .Font.Color.RGB = PRIMARY_RED
    End With
    
    With sld.Shapes(2).TextFrame.TextRange
        .text = "Project Status Report" & vbCrLf & Format(Date, "mmmm yyyy")
        .Font.Name = "Arial"
        .Font.Size = 24
        .Font.Color.RGB = SECONDARY_GREY
    End With
    
    Set shp = sld.Shapes.AddLine(100, 330, 620, 330)
    shp.Line.ForeColor.RGB = PRIMARY_RED
    shp.Line.Weight = 3
    
    ' ----- SLIDE 2: PROJECT OVERVIEW -----
    slideIndex = slideIndex + 1
    Set sld = ppt.Slides.Add(slideIndex, 11)
    
    With sld.Shapes(1).TextFrame.TextRange
        .text = "Project Overview"
        .Font.Name = "Arial"
        .Font.Bold = True
        .Font.Color.RGB = PRIMARY_RED
    End With
    
    ' Big completion percentage
    Set shp = sld.Shapes.AddTextbox(1, 480, 100, 200, 100)
    With shp.TextFrame.TextRange
        .text = Format(completionRate, "0") & "%"
        .Font.Name = "Arial"
        .Font.Size = 54
        .Font.Bold = True
        .Font.Color.RGB = PRIMARY_RED
        .ParagraphFormat.Alignment = 2
    End With
    
    Set shp = sld.Shapes.AddTextbox(1, 480, 195, 200, 30)
    With shp.TextFrame.TextRange
        .text = "Complete"
        .Font.Name = "Arial"
        .Font.Size = 16
        .Font.Color.RGB = SECONDARY_GREY
        .ParagraphFormat.Alignment = 2
    End With
    
    ' Project details table
    Set shp = sld.Shapes.AddTable(6, 2, 40, 110, 300, 160)
    Set tbl = shp.Table
    
    tbl.Cell(1, 1).Shape.TextFrame.TextRange.text = "Status"
    tbl.Cell(1, 2).Shape.TextFrame.TextRange.text = wsProject.Range("B6").Value
    tbl.Cell(2, 1).Shape.TextFrame.TextRange.text = "Priority"
    tbl.Cell(2, 2).Shape.TextFrame.TextRange.text = wsProject.Range("B7").Value
    tbl.Cell(3, 1).Shape.TextFrame.TextRange.text = "Start Date"
    tbl.Cell(3, 2).Shape.TextFrame.TextRange.text = wsProject.Range("B8").Value
    tbl.Cell(4, 1).Shape.TextFrame.TextRange.text = "End Date"
    tbl.Cell(4, 2).Shape.TextFrame.TextRange.text = wsProject.Range("B9").Value
    tbl.Cell(5, 1).Shape.TextFrame.TextRange.text = "Total Tasks"
    tbl.Cell(5, 2).Shape.TextFrame.TextRange.text = CStr(taskCount)
    tbl.Cell(6, 1).Shape.TextFrame.TextRange.text = "Closed Tasks"
    tbl.Cell(6, 2).Shape.TextFrame.TextRange.text = CStr(closedCount)
    
    For i = 1 To 6
        tbl.Cell(i, 1).Shape.Fill.ForeColor.RGB = PRIMARY_RED
        tbl.Cell(i, 1).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        tbl.Cell(i, 1).Shape.TextFrame.TextRange.Font.Bold = True
        tbl.Cell(i, 1).Shape.TextFrame.TextRange.Font.Name = "Arial"
        tbl.Cell(i, 1).Shape.TextFrame.TextRange.Font.Size = 11
        tbl.Cell(i, 2).Shape.TextFrame.TextRange.Font.Name = "Arial"
        tbl.Cell(i, 2).Shape.TextFrame.TextRange.Font.Size = 11
    Next i
    
    ' Status breakdown table
    Set shp = sld.Shapes.AddTable(4, 2, 40, 290, 200, 110)
    Set tbl = shp.Table
    
    tbl.Cell(1, 1).Shape.TextFrame.TextRange.text = "Task Status"
    tbl.Cell(1, 2).Shape.TextFrame.TextRange.text = "Count"
    FormatPPTTableHeader tbl, 1, 2
    
    tbl.Cell(2, 1).Shape.TextFrame.TextRange.text = "Closed"
    tbl.Cell(2, 1).Shape.Fill.ForeColor.RGB = STATUS_GREEN
    tbl.Cell(2, 2).Shape.TextFrame.TextRange.text = CStr(closedCount)
    
    tbl.Cell(3, 1).Shape.TextFrame.TextRange.text = "In Progress"
    tbl.Cell(3, 1).Shape.Fill.ForeColor.RGB = STATUS_YELLOW
    tbl.Cell(3, 2).Shape.TextFrame.TextRange.text = CStr(inProgressCount)
    
    tbl.Cell(4, 1).Shape.TextFrame.TextRange.text = "Open"
    tbl.Cell(4, 1).Shape.Fill.ForeColor.RGB = LIGHT_GREY
    tbl.Cell(4, 2).Shape.TextFrame.TextRange.text = CStr(openCount)
    
    For i = 2 To 4
        For j = 1 To 2
            tbl.Cell(i, j).Shape.TextFrame.TextRange.Font.Name = "Arial"
            tbl.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = 11
        Next j
    Next i
    
    ' ----- SLIDE 3: ACTIVE TASKS (Open & In Progress Only) -----
    slideIndex = slideIndex + 1
    Set sld = ppt.Slides.Add(slideIndex, 11)
    
    With sld.Shapes(1).TextFrame.TextRange
        .text = "Active Tasks"
        .Font.Name = "Arial"
        .Font.Bold = True
        .Font.Color.RGB = PRIMARY_RED
    End With
    
    ' Count active tasks (Open + In Progress only)
    Dim activeTaskCount As Long
    Dim lastTaskRow As Long
    lastTaskRow = wsTasks.Cells(wsTasks.Rows.count, "A").End(xlUp).row
    
    activeTaskCount = 0
    For r = 4 To lastTaskRow
        If wsTasks.Cells(r, 4).Value = "Open" Or wsTasks.Cells(r, 4).Value = "In Progress" Then
            activeTaskCount = activeTaskCount + 1
        End If
    Next r
    
    Dim taskTableRows As Long
    taskTableRows = activeTaskCount + 1
    If taskTableRows > 15 Then taskTableRows = 15
    If taskTableRows < 2 Then taskTableRows = 2 ' At least header + 1 row
    
    If activeTaskCount > 0 Then
        Set shp = sld.Shapes.AddTable(taskTableRows, 4, 30, 95, 660, 26 * taskTableRows)
        Set tbl = shp.Table
        
        tbl.Cell(1, 1).Shape.TextFrame.TextRange.text = "Task"
        tbl.Cell(1, 2).Shape.TextFrame.TextRange.text = "Status"
        tbl.Cell(1, 3).Shape.TextFrame.TextRange.text = "Duration"
        tbl.Cell(1, 4).Shape.TextFrame.TextRange.text = "Due Date"
        FormatPPTTableHeader tbl, 1, 4
        
        Dim tableRow As Long
        tableRow = 2
        
        For r = 4 To lastTaskRow
            ' Only include Open and In Progress tasks
            If wsTasks.Cells(r, 4).Value = "Open" Or wsTasks.Cells(r, 4).Value = "In Progress" Then
                If tableRow <= taskTableRows Then
                    tbl.Cell(tableRow, 1).Shape.TextFrame.TextRange.text = Left(CStr(wsTasks.Cells(r, 2).Value), 40)
                    tbl.Cell(tableRow, 2).Shape.TextFrame.TextRange.text = CStr(wsTasks.Cells(r, 4).Value)
                    tbl.Cell(tableRow, 3).Shape.TextFrame.TextRange.text = CStr(wsTasks.Cells(r, 8).Value)
                    tbl.Cell(tableRow, 4).Shape.TextFrame.TextRange.text = CStr(wsTasks.Cells(r, 7).Value)
                    
                    ' Status colors
                    Select Case wsTasks.Cells(r, 4).Value
                        Case "In Progress": tbl.Cell(tableRow, 2).Shape.Fill.ForeColor.RGB = STATUS_YELLOW
                        Case "Open": tbl.Cell(tableRow, 2).Shape.Fill.ForeColor.RGB = LIGHT_GREY
                    End Select
                    
                    For j = 1 To 4
                        tbl.Cell(tableRow, j).Shape.TextFrame.TextRange.Font.Name = "Arial"
                        tbl.Cell(tableRow, j).Shape.TextFrame.TextRange.Font.Size = 9
                    Next j
                    
                    ' Alternate row shading
                    If tableRow Mod 2 = 0 Then
                        For j = 1 To 4
                            If j <> 2 Then tbl.Cell(tableRow, j).Shape.Fill.ForeColor.RGB = RGB(248, 248, 248)
                        Next j
                    End If
                    
                    tableRow = tableRow + 1
                End If
            End If
        Next r
        
        ' Note if more active tasks
        If activeTaskCount > 14 Then
            Set shp = sld.Shapes.AddTextbox(1, 30, 480, 660, 25)
            With shp.TextFrame.TextRange
                .text = "... and " & (activeTaskCount - 14) & " more active tasks (see Excel for full list)"
                .Font.Name = "Arial"
                .Font.Size = 10
                .Font.Italic = True
                .Font.Color.RGB = SECONDARY_GREY
            End With
        End If
    Else
        ' No active tasks - show message
        Set shp = sld.Shapes.AddTextbox(1, 30, 150, 660, 50)
        With shp.TextFrame.TextRange
            .text = "All tasks are complete!"
            .Font.Name = "Arial"
            .Font.Size = 24
            .Font.Color.RGB = STATUS_GREEN
            .ParagraphFormat.Alignment = 2
        End With
    End If
    
    ' ----- SLIDE 4: WHY WE TRACK -----
    slideIndex = slideIndex + 1
    Set sld = ppt.Slides.Add(slideIndex, 11)
    
    With sld.Shapes(1).TextFrame.TextRange
        .text = "Why We Track"
        .Font.Name = "Arial"
        .Font.Bold = True
        .Font.Color.RGB = PRIMARY_RED
    End With
    
    Set shp = sld.Shapes.AddTextbox(1, 40, 100, 640, 300)
    With shp.TextFrame.TextRange
        .text = "1. VISIBILITY PREVENTS SURPRISES" & vbCrLf & _
                "    Leadership should never learn about issues from end users." & vbCrLf & vbCrLf & _
                "2. DEPENDENCIES HIDE IN PLAIN SIGHT" & vbCrLf & _
                "    Projects touch multiple teams. Tracking surfaces connections." & vbCrLf & vbCrLf & _
                "3. FUTURE PLANNING NEEDS CURRENT CLARITY" & vbCrLf & _
                "    You can't commit to new work if current projects slip." & vbCrLf & vbCrLf & _
                "4. DOCUMENTATION PROTECTS CONTINUITY" & vbCrLf & _
                "    A maintained register means work survives staff changes." & vbCrLf & vbCrLf & _
                "5. PATTERNS EMERGE OVER TIME" & vbCrLf & _
                "    Tracking reveals systemic issues and drives improvement."
        .Font.Name = "Arial"
        .Font.Size = 13
        .Font.Color.RGB = SECONDARY_GREY
    End With
    
    Set shp = sld.Shapes.AddLine(40, 420, 680, 420)
    shp.Line.ForeColor.RGB = PRIMARY_RED
    shp.Line.Weight = 2
    
    Application.StatusBar = False
    
    MsgBox "PowerPoint created!" & vbCrLf & vbCrLf & _
           "Project: " & projectName & vbCrLf & _
           "Slides: " & slideIndex, vbInformation, "Export Complete"
    
End Sub

Sub FormatPPTTableHeader(tbl As Object, rowNum As Long, colCount As Long)
    Dim j As Long
    For j = 1 To colCount
        tbl.Cell(rowNum, j).Shape.Fill.ForeColor.RGB = PRIMARY_RED
        tbl.Cell(rowNum, j).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        tbl.Cell(rowNum, j).Shape.TextFrame.TextRange.Font.Bold = True
        tbl.Cell(rowNum, j).Shape.TextFrame.TextRange.Font.Name = "Arial"
        tbl.Cell(rowNum, j).Shape.TextFrame.TextRange.Font.Size = 10
    Next j
End Sub

' ============================================
' STATUS AND TYPE HELPER FUNCTIONS
' ============================================
Function GetTaskStatusName(statusId As String) As String
    Select Case statusId
        Case STATUS_OPEN_ID: GetTaskStatusName = "Open"
        Case STATUS_IN_PROGRESS_ID: GetTaskStatusName = "In Progress"
        Case STATUS_CLOSED_ID: GetTaskStatusName = "Closed"
        Case Else: GetTaskStatusName = "Unknown"
    End Select
End Function

Function GetTaskStatusColor(statusId As String) As Long
    Select Case statusId
        Case STATUS_OPEN_ID: GetTaskStatusColor = LIGHT_GREY
        Case STATUS_IN_PROGRESS_ID: GetTaskStatusColor = STATUS_YELLOW
        Case STATUS_CLOSED_ID: GetTaskStatusColor = STATUS_GREEN
        Case Else: GetTaskStatusColor = LIGHT_GREY
    End Select
End Function

Function GetTypeName(typeId As String) As String
    Select Case typeId
        Case TYPE_FOLDER_ID: GetTypeName = "Folder"
        Case TYPE_PROJECT_ID: GetTypeName = "Project"
        Case TYPE_TASK_ID: GetTypeName = "Task"
        Case Else: GetTypeName = "Unknown"
    End Select
End Function

Function GetStatusName(statusId As String) As String
    ' Project-level status - try task IDs first, then fall back to simple 1-4
    Select Case statusId
        Case STATUS_OPEN_ID, "1": GetStatusName = "Open"
        Case STATUS_IN_PROGRESS_ID, "2": GetStatusName = "In Progress"
        Case STATUS_CLOSED_ID, "3": GetStatusName = "Closed"
        Case "4": GetStatusName = "On Hold"
        Case Else: GetStatusName = "Unknown"
    End Select
End Function

Function GetStatusColor(statusId As String) As Long
    Select Case statusId
        Case STATUS_OPEN_ID, "1": GetStatusColor = LIGHT_GREY
        Case STATUS_IN_PROGRESS_ID, "2": GetStatusColor = STATUS_YELLOW
        Case STATUS_CLOSED_ID, "3": GetStatusColor = STATUS_GREEN
        Case "4": GetStatusColor = STATUS_RED
        Case Else: GetStatusColor = LIGHT_GREY
    End Select
End Function

Function GetPriorityName(priorityId As String) As String
    Select Case priorityId
        Case "1", "5000077591": GetPriorityName = "Low"
        Case "2", "5000077592": GetPriorityName = "Medium"
        Case "3", "5000077593": GetPriorityName = "High"
        Case "4", "5000077594": GetPriorityName = "Urgent"
        Case Else: GetPriorityName = "None"
    End Select
End Function

' ============================================
' API CALL FUNCTION
' ============================================
Function MakeAPICall(url As String, apiKey As String) As String
    On Error GoTo ErrHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    Dim authString As String
    authString = Base64Encode(apiKey & ":X")
    
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Basic " & authString
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.send
    
    If http.status = 200 Then
        MakeAPICall = http.responseText
    Else
        MakeAPICall = "ERROR " & http.status & ": " & http.statusText
    End If
    Exit Function
    
ErrHandler:
    MakeAPICall = "ERROR: " & Err.Description
End Function

Function Base64Encode(text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
    Dim objXML As Object, objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = Replace(objNode.text, vbLf, "")
End Function

' ============================================
' JSON PARSING HELPERS
' ============================================
Function ExtractString(json As String, startPos As Long, key As String) As String
    Dim keyPos As Long
    keyPos = InStr(startPos, json, """" & key & """:""")
    If keyPos = 0 Or keyPos > startPos + 2000 Then Exit Function
    Dim valStart As Long: valStart = keyPos + Len(key) + 4
    Dim valEnd As Long: valEnd = InStr(valStart, json, """")
    If valEnd > valStart Then ExtractString = Mid(json, valStart, valEnd - valStart)
End Function

Function ExtractNumber(json As String, startPos As Long) As String
    If startPos < 1 Then Exit Function
    Dim i As Long, c As String, result As String
    For i = startPos To startPos + 15
        c = Mid(json, i, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
        ElseIf result <> "" Then
            Exit For
        End If
    Next i
    ExtractNumber = result
End Function

' ============================================
' DATE FORMATTING
' ============================================
Function FormatAPIDate(dateStr As String) As String
    If dateStr = "" Or dateStr = "null" Then
        FormatAPIDate = ""
        Exit Function
    End If
    If Len(dateStr) >= 10 Then
        FormatAPIDate = Left(dateStr, 10)
    Else
        FormatAPIDate = dateStr
    End If
End Function

' ============================================
' SHEET HELPER
' ============================================
Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
    On Error GoTo 0
End Function

Sub FormatHeader(rng As Range)
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = PRIMARY_RED
    rng.Font.Name = "Arial"
End Sub

