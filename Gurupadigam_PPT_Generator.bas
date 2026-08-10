Option Explicit

' ===== PowerPoint constants =====
Private Const ppLayoutBlank As Long = 12
Private Const ppAlignJustify As Long = 4
Private Const msoTextOrientationHorizontal As Long = 1
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0

' ===== CM ? Points =====
Private Function CmToPoints(cm As Double) As Double
    CmToPoints = cm * 28.35
End Function

Sub CreatePPTReports()

    Dim pptApp As Object, ppt As Object, sld As Object, tb As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Randomize

    ' ===== Excel Sheet =====
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows found.", vbExclamation
        Exit Sub
    End If

    ' ===== Open PowerPoint =====
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0

    pptApp.Visible = True
    Set ppt = pptApp.Presentations.Add

    ' ===== Photo Folder =====
    Dim photoPath As String
    photoPath = "F:\gurupadigam_package\photos\"
    If Right(photoPath, 1) <> "\" Then photoPath = photoPath & "\"

    ' ===== Loop students =====
    For i = 2 To lastRow

        Dim regNo As String, menteeName As String, slotTxt As String
        Dim courseCode As String, courseName As String, interest As String
        Dim attendance As Double, maxMarks As Double, testMarks As Double
        Dim classAvg As Variant
        Dim para As String, para2 As String, fullText As String

        regNo = CStr(ws.Cells(i, 1).Value)
        menteeName = CStr(ws.Cells(i, 2).Value)
        interest = CStr(ws.Cells(i, 3).Value)
        courseCode = CStr(ws.Cells(i, 4).Value)
        courseName = CStr(ws.Cells(i, 5).Value)
        attendance = ParseToDouble(ws.Cells(i, 6).Value)
        slotTxt = CStr(ws.Cells(i, 7).Value)
        maxMarks = ParseToDouble(ws.Cells(i, 8).Value)
        testMarks = ParseToDouble(ws.Cells(i, 9).Value)
        classAvg = ws.Cells(i, 10).Value

        ' =========================
        ' SLIDE 1 - DETAILS
        ' =========================
        Set sld = ppt.Slides.Add(ppt.Slides.Count + 1, ppLayoutBlank)

        ' --- Table 1: Name / Reg ---
        Dim t1 As Object
        Set t1 = sld.Shapes.AddTable(2, 2, _
            CmToPoints(1.8), CmToPoints(3.5), CmToPoints(14), CmToPoints(3)).Table

        t1.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Name"
        t1.Cell(1, 2).Shape.TextFrame.TextRange.Text = menteeName
        t1.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Reg No"
        t1.Cell(2, 2).Shape.TextFrame.TextRange.Text = regNo
        FormatTable t1, "Calibri", 20

        ' --- Photo ---
        Dim photoFile As String
        photoFile = FindPhotoFile(photoPath, regNo)
        If photoFile <> "" Then
            sld.Shapes.AddPicture photoFile, msoFalse, msoTrue, _
                CmToPoints(1.8), CmToPoints(7), CmToPoints(7), CmToPoints(7)
        End If

        ' --- Table 2: Course details ---
        Dim t2 As Object
        Set t2 = sld.Shapes.AddTable(2, 4, _
            CmToPoints(13.5), CmToPoints(3.5), CmToPoints(20), CmToPoints(4)).Table

        t2.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Slot"
        t2.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Course Code"
        t2.Cell(1, 3).Shape.TextFrame.TextRange.Text = "Course Name"
        t2.Cell(1, 4).Shape.TextFrame.TextRange.Text = "Attendance %"

        t2.Cell(2, 1).Shape.TextFrame.TextRange.Text = slotTxt
        t2.Cell(2, 2).Shape.TextFrame.TextRange.Text = courseCode
        t2.Cell(2, 3).Shape.TextFrame.TextRange.Text = courseName
        t2.Cell(2, 4).Shape.TextFrame.TextRange.Text = IIf(attendance > 0, attendance & "%", "")
        FormatTable t2, "Calibri", 20

        ' --- Table 3: Test marks ---
        Dim t3 As Object
        Set t3 = sld.Shapes.AddTable(2, 4, _
            CmToPoints(13.5), CmToPoints(13.0), CmToPoints(20), CmToPoints(4)).Table

        t3.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Slot"
        t3.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Test Marks"
        t3.Cell(1, 3).Shape.TextFrame.TextRange.Text = "Maximum Marks"
        t3.Cell(1, 4).Shape.TextFrame.TextRange.Text = "Class Average"

        t3.Cell(2, 1).Shape.TextFrame.TextRange.Text = slotTxt
        t3.Cell(2, 2).Shape.TextFrame.TextRange.Text = SafeDisplay(ws.Cells(i, 9).Value)
        t3.Cell(2, 3).Shape.TextFrame.TextRange.Text = SafeDisplay(ws.Cells(i, 8).Value)
        t3.Cell(2, 4).Shape.TextFrame.TextRange.Text = SafeDisplay(classAvg)
        FormatTable t3, "Calibri", 20

        ' =========================
        ' SLIDE 2 - VARIED MESSAGE
        ' =========================
        Set sld = ppt.Slides.Add(ppt.Slides.Count + 1, ppLayoutBlank)

        para = "Dear Parents," & vbCrLf & vbCrLf & PickOne(Array( _
               "Your son Mr " & menteeName & " is pursuing the course '" & courseName & "' in our institution. ", _
               "We wish to inform you about the academic progress of your son Mr " & menteeName & ", who is enrolled in '" & courseName & "'. ", _
               "Mr " & menteeName & " is currently continuing his studies in the course '" & courseName & "'. "))

        para = para & "His current attendance is " & _
               IIf(attendance > 0, attendance & "%. ", "not available. ")

        If attendance > 0 And attendance < 80 Then
            para = para & PickOne(Array( _
                "Regular attendance is essential, and improvement is required in this regard. ", _
                "We request your support in motivating him to attend classes consistently. "))
        ElseIf attendance >= 80 Then
            para = para & PickOne(Array( _
                "His attendance record is satisfactory and reflects good academic discipline. ", _
                "He has been regular to classes, which is appreciated. "))
        End If

        If maxMarks > 0 Then
            If testMarks / maxMarks < 0.6 Then
                para = para & PickOne(Array( _
                    "His internal assessment performance indicates scope for improvement. ", _
                    "Additional focus and structured preparation are advised to enhance performance. "))
            Else
                para = para & PickOne(Array( _
                    "His internal assessment performance is satisfactory. ", _
                    "The test results indicate an acceptable level of understanding. "))
            End If
        End If

        para = para & PickOne(Array( _
            "Students are encouraged to enroll in NPTEL courses to strengthen subject knowledge. ", _
            "Participation in NPTEL and other recognized online certification courses is highly recommended. "))

        para = para & PickOne(Array( _
            "Online learning platforms can help develop technical and analytical skills. ", _
            "Skill-based online courses will support both academic and career development. "))

        If Len(Trim(interest)) > 0 Then
            para2 = "He is also encouraged to pursue his interest in " & Trim(interest) & _
                    " in a balanced manner alongside academics. "
        Else
            para2 = ""
        End If

        para2 = para2 & PickOne(Array( _
            "Students must strictly adhere to college rules regarding uniform, grooming, and classroom discipline. ", _
            "Maintaining discipline and following institutional norms are mandatory for all students. ")) & _
            "Parental guidance and regular monitoring will greatly support his academic progress."

        fullText = para & vbCrLf & vbCrLf & para2

Set tb = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
    CmToPoints(3), CmToPoints(), CmToPoints(28), CmToPoints(12))

' White background for textbox
With tb.Fill
    .Visible = msoTrue
    .Solid
    .ForeColor.RGB = RGB(255, 255, 255)
End With

' Hide border
tb.Line.Visible = msoFalse

tb.TextFrame.TextRange.Text = fullText

With tb.TextFrame.TextRange
    .Font.Name = "Calibri"
    .Font.Size = 20
    .ParagraphFormat.Alignment = ppAlignJustify
End With
        BoldText tb.TextFrame.TextRange, "Mr " & menteeName
        BoldText tb.TextFrame.TextRange, courseName

    Next i

    ppt.SaveAs ThisWorkbook.Path & "\Gurupadigam_Output.pptx"
    MsgBox "PPT created successfully.", vbInformation

End Sub

' ===== Helper procedures =====
Private Function PickOne(arr As Variant) As String
    PickOne = arr(Int(Rnd * (UBound(arr) + 1)))
End Function

Private Function ParseToDouble(v As Variant) As Double
    Dim s As String
    s = Replace(Replace(Trim(CStr(v)), "%", ""), ",", "")
    If IsNumeric(s) Then ParseToDouble = CDbl(s)
End Function

Private Function SafeDisplay(v As Variant) As String
    SafeDisplay = Trim(CStr(v))
End Function

Private Sub FormatTable(tbl As Object, fontName As String, fontSize As Long)
    Dim r As Long, c As Long
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            With tbl.Cell(r, c).Shape.TextFrame.TextRange
                .Font.Name = fontName
                .Font.Size = fontSize
            End With
        Next c
    Next r
End Sub

Private Sub BoldText(tr As Object, searchText As String)
    Dim f As Object
    Set f = tr.Find(searchText)
    If Not f Is Nothing Then f.Font.Bold = True
End Sub

Private Function FindPhotoFile(folderPath As String, regNo As String) As String
    Dim ext As Variant, f As String
    For Each ext In Array("jpg", "jpeg", "png", "tif")
        f = Dir(folderPath & regNo & "." & ext)
        If f <> "" Then
            FindPhotoFile = folderPath & f
            Exit Function
        End If
    Next ext
    FindPhotoFile = ""
End Function