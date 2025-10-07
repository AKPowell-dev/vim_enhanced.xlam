Attribute VB_Name = "Z_Auxiliary"
Option Explicit
Option Private Module
' new function
Function CycleFillColor(Optional ByVal g As String) As Boolean
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim colors As Variant
    Static lastIndex As Long
    Static lastAddress As String

    colors = Array(xlNone, RGB(0, 32, 96), RGB(220, 228, 244), _
                   RGB(240, 240, 240), RGB(255, 242, 204))

    ' Check selection
    If TypeName(Selection) <> "Range" Then GoTo CleanExit

    ' Reset index if new selection
    If Selection.Address <> lastAddress Then
        lastIndex = 0
        lastAddress = Selection.Address
    End If

    ' Cycle to next color
    lastIndex = lastIndex + 1
    If lastIndex > UBound(colors) Then lastIndex = 0

    ' Apply color
    If colors(lastIndex) = xlNone Then
        Selection.Interior.ColorIndex = xlNone
    Else
        Selection.Interior.Color = colors(lastIndex)
    End If

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Function

Sub CycleFontColor()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim fontColorsArray As Variant
    Static fontCycleIndex As Long
    Static fontCycleLastAddress As String

    ' Exact colors
    fontColorsArray = Array(RGB(0, 0, 0), RGB(255, 255, 255), RGB(0, 0, 255), RGB(153, 0, 0), RGB(0, 128, 0))

    ' Exit if selection is not a range
    If TypeName(Selection) <> "Range" Then GoTo CleanExit

    ' Reset cycle if selection changes
    If Selection.Address <> fontCycleLastAddress Then
        fontCycleIndex = 0
        fontCycleLastAddress = Selection.Address
    End If

    ' Apply color
    Selection.Font.Color = fontColorsArray(fontCycleIndex)

    ' Advance index
    fontCycleIndex = fontCycleIndex + 1
    If fontCycleIndex > UBound(fontColorsArray) Then fontCycleIndex = 0

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

Function CycleNumberFormat(Optional ByVal g As String) As Boolean
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim formats As Variant
    Static lastIndex As Long
    Static lastAddress As String
    
    formats = Array( _
        "#,##0_);(#,##0);--_)", _
        "$#,##0_);($#,##0);$--_)", _
        "#,##0.0%_);(#,##0.0%);--\%_)", _
        "#,##0.0x_);(#,##0.0x);--x_)", _
        "#,##0""bps""_);(#,##0""bps"");""--bps """, _
        """On"";"""";""Off""", _
        "[=1]""Yes"";[=0]""No"";""ERROR""", _
        "[=1]0"" Year"";0"" Years""", _
        """Year ""0; ""Year ""-0; ""Year 0""; """"" _
    )
    
    If TypeName(Selection) <> "Range" Then
        CycleNumberFormat = False
        Exit Function
    End If
    
    ' Reset if new selection
    If Selection.Address <> lastAddress Then
        lastIndex = 0
        lastAddress = Selection.Address
    End If
    
    ' Apply to entire selection at once
    Application.ScreenUpdating = False
    Selection.NumberFormat = formats(lastIndex)
    Application.ScreenUpdating = True
    
    ' Advance the cycle
    lastIndex = lastIndex + 1
    If lastIndex > UBound(formats) Then lastIndex = 0
    
    CycleNumberFormat = False
    
' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Function


Public Sub DeleteLikeExcel()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    If TypeName(Selection) <> "Range" Then Exit Sub
    ' Send the Delete key to Excel
    Application.SendKeys "{DEL}", True
    
' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

Sub ClearFormatting()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then Exit Sub

    ' Speed up
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    On Error Resume Next
    With rng
        .Borders.LineStyle = xlNone
        .Interior.ColorIndex = xlNone
        .NumberFormat = "#,##0_);(#,##0);--_)"
        ' Remove bold and italic for the whole selection in one pass
        .Font.Bold = False
        .Font.Italic = False
    End With
    On Error GoTo 0

    ' Restore
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

Sub CycleRowHeight()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range
    Dim currentHeight As Double
    Dim nextHeight As Double
    
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    Set sel = Selection
    
    currentHeight = sel.Rows(1).RowHeight
    
    ' Decide next height
    If Abs(currentHeight - 3) < 0.1 Then
        nextHeight = 15
    Else
        nextHeight = 3
    End If
    
    ' Apply to all selected rows
    sel.EntireRow.RowHeight = nextHeight

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub


Sub CycleColumnWidth()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range
    Dim currentWidth As Double
    Dim nextWidth As Double
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    
    ' Use the first column in selection to determine current width
    currentWidth = sel.Columns(1).ColumnWidth
    
    ' Decide next width
    If Abs(currentWidth - 8.43) < 0.1 Then
        nextWidth = 0.5
    Else
        nextWidth = 8.43
    End If
    
    ' Apply width to all selected columns
    sel.EntireColumn.ColumnWidth = nextWidth
    
' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub


Sub CycleBorder(side As String)
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim borderStyles As Variant
    Static lastIndex As Long
    Static lastAddress As String
    Dim rng As Range
    Dim b As Border

    ' Only run on range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection

    ' Reset cycle if selection changes
    If rng.Address <> lastAddress Then
        lastIndex = 0
        lastAddress = rng.Address
    End If

    ' Border styles to cycle: continuous, double, dotted, none
    borderStyles = Array(xlNone, xlDot, xlDouble)

    ' Advance cycle
    lastIndex = lastIndex + 1
    If lastIndex > UBound(borderStyles) Then lastIndex = 0

    ' Apply to the specified side
    Select Case LCase(side)
        Case "h": Set b = rng.Borders(xlEdgeLeft)
        Case "j": Set b = rng.Borders(xlEdgeBottom)
        Case "k": Set b = rng.Borders(xlEdgeTop)
        Case "l": Set b = rng.Borders(xlEdgeRight)
        Case Else: Exit Sub
    End Select

    b.LineStyle = borderStyles(lastIndex)
    
' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub
Sub CycleBorderLeft(): CycleBorder "h": End Sub
Sub CycleBorderBottom(): CycleBorder "j": End Sub
Sub CycleBorderTop(): CycleBorder "k": End Sub
Sub CycleBorderRight(): CycleBorder "l": End Sub

Sub CopyPasteAsPictureToPPT()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Const msoTrue As Long = -1
    Dim pptApp As Object, pptSlide As Object
    Dim targetShape As Object, pastedShp As Object
    Dim selRange As Range
    
    ' Ensure selection is valid
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        Exit Sub
    End If
    Set selRange = Selection
    
    ' Copy as picture (as shown when printed)
    selRange.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
    
    ' Connect to running PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        MsgBox "PowerPoint is not running.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get active slide
    Set pptSlide = pptApp.ActiveWindow.View.Slide
    
    ' Check for currently selected shape to replace
    On Error Resume Next
    Set targetShape = pptApp.ActiveWindow.Selection.ShapeRange
    On Error GoTo 0
    
    ' Paste the picture
    Set pastedShp = pptSlide.Shapes.Paste(1) ' ppPasteEnhancedMetafile
    
    ' If a target shape exists, resize & overlay it
    If Not targetShape Is Nothing Then
        With pastedShp
            .Left = targetShape.Left
            .Top = targetShape.Top
            .Width = targetShape.Width
            .Height = targetShape.Height
        End With
        targetShape.Delete
    End If
    
    ' Center the pasted shape if no target
    If targetShape Is Nothing Then
        With pastedShp
            .Left = (pptSlide.Master.Width - .Width) / 2
            .Top = (pptSlide.Master.Height - .Height) / 2
        End With
    End If
    
    ' Optional: select the new picture
    pastedShp.Select

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

' Vimium mapping wrapper
Sub PasteAsPicture_XP(): CopyPasteAsPictureToPPT_XP: End Sub

Sub SmartFillRight()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range, ws As Worksheet
    Dim r As Long, startRow As Long, startCol As Long, lastCol As Long
    Dim headerRows As Long, i As Long, c As Long
    Dim sourceCell As Range
    Dim sourceHasFill As Boolean, sourceFillColor As Long
    Dim belowHasText As Boolean
    Dim overallLastCol As Long
    Dim resultRange As Range

    headerRows = 15
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    Set sel = Selection
    Set ws = sel.Worksheet

    startRow = sel.Row
    startCol = sel.Column
    overallLastCol = startCol

    For r = 1 To sel.Rows.Count
        Set sourceCell = sel.Cells(r, 1)
        sourceHasFill = (sourceCell.Interior.ColorIndex <> xlNone)
        sourceFillColor = sourceCell.Interior.Color

        ' Determine last column to fill for this row
        lastCol = startCol
        For c = startCol + 1 To ws.Columns.Count
            Dim headerHasText As Boolean
            headerHasText = False

            ' Check up to headerRows above safely
            For i = 1 To headerRows
                If (startRow - i) >= 1 Then
                    If ws.Cells(startRow - i, c).Value <> "" Then
                        headerHasText = True
                        Exit For
                    End If
                End If
            Next i

            ' Check up to headerRows below safely
            belowHasText = False
            For i = 1 To headerRows
                If (startRow + r - 1 + i) <= ws.Rows.Count Then
                    If ws.Cells(startRow + r - 1 + i, c).Value <> "" Then
                        belowHasText = True
                        Exit For
                    End If
                End If
            Next i

            If headerHasText Or belowHasText Then
                lastCol = c
            Else
                Exit For
            End If
        Next c

        ' Fill entire row dynamically and copy formatting
        If lastCol > startCol Then
            ws.Range(ws.Cells(startRow + r - 1, startCol), ws.Cells(startRow + r - 1, lastCol)).FillRight
            With ws.Range(ws.Cells(startRow + r - 1, startCol + 1), ws.Cells(startRow + r - 1, lastCol))
                .Font.Name = sourceCell.Font.Name
                .Font.Size = sourceCell.Font.Size
                .Font.Bold = sourceCell.Font.Bold
                .Font.Italic = sourceCell.Font.Italic
                .NumberFormat = sourceCell.NumberFormat
                If sourceHasFill Then
                    .Interior.Color = sourceFillColor
                Else
                    .Interior.Pattern = xlNone
                End If
            End With

            If lastCol > overallLastCol Then overallLastCol = lastCol
        End If
    Next r

    ' Select the full operated block (rows x from startCol to overallLastCol)
    If overallLastCol > startCol Then
        Set resultRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + sel.Rows.Count - 1, overallLastCol))
        resultRange.Select
    Else
        sel.Select
    End If

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

Sub SmartFillDown()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range, ws As Worksheet
    Dim c As Long, startRow As Long, startCol As Long, lastRow As Long
    Dim headerCols As Long, i As Long, r As Long
    Dim sourceCell As Range, originalCell As Range
    Dim sourceHasFill As Boolean, sourceFillColor As Long
    Dim leftHasText As Boolean, rightHasText As Boolean
    
    headerCols = 5 ' check up to 5 columns left/right
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    Set ws = sel.Worksheet
    Set originalCell = ActiveCell
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    startRow = sel.Row
    startCol = sel.Column
    
    For c = 1 To sel.Columns.Count
        Set sourceCell = sel.Cells(1, c)
        sourceHasFill = (sourceCell.Interior.ColorIndex <> xlNone)
        sourceFillColor = sourceCell.Interior.Color
        
        ' Determine last row to fill
        lastRow = startRow
        For r = startRow + 1 To ws.Rows.Count
            leftHasText = False
            rightHasText = False
            
            ' Check up to headerCols to the left
            For i = 1 To headerCols
                If (startCol + c - 1 - i) >= 1 Then
                    If ws.Cells(r, startCol + c - 1 - i).Value <> "" Then
                        leftHasText = True
                        Exit For
                    End If
                End If
            Next i
            
            ' Check up to headerCols to the right
            For i = 1 To headerCols
                If (startCol + c - 1 + i) <= ws.Columns.Count Then
                    If ws.Cells(r, startCol + c - 1 + i).Value <> "" Then
                        rightHasText = True
                        Exit For
                    End If
                End If
            Next i
            
            ' Extend lastRow if either left or right has text
            If leftHasText Or rightHasText Then
                lastRow = r
            Else
                Exit For
            End If
        Next r
        
        ' Fill entire column dynamically
        If lastRow > startRow Then
            ws.Range(ws.Cells(startRow, startCol + c - 1), ws.Cells(lastRow, startCol + c - 1)).FillDown
            ' Copy source formatting
            With ws.Range(ws.Cells(startRow + 1, startCol + c - 1), ws.Cells(lastRow, startCol + c - 1))
                .Font.Name = sourceCell.Font.Name
                .Font.Size = sourceCell.Font.Size
                .Font.Bold = sourceCell.Font.Bold
                .Font.Italic = sourceCell.Font.Italic
                .NumberFormat = sourceCell.NumberFormat
                
                If sourceHasFill Then
                    .Interior.Color = sourceFillColor
                Else
                    .Interior.Pattern = xlNone
                End If
            End With
        End If
    Next c
    
    originalCell.Select
    
' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

Sub CenterAcrossSelection()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range
    
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    Set sel = Selection
    
    ' Apply "Center Across Selection" alignment
    With sel
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub


Sub WrapInIFERROR()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range
    Dim formulas As Variant
    Dim r As Long, c As Long

    ' Ensure something is selected
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    Set sel = Selection

    ' Read all formulas into an array
    formulas = sel.Formula

    ' Loop through array, wrap formulas with IFERROR
    For r = 1 To UBound(formulas, 1)
        For c = 1 To UBound(formulas, 2)
            If Left(formulas(r, c), 1) = "=" Then
                formulas(r, c) = "=IFERROR(" & Mid(formulas(r, c), 2) & ",0)"
            End If
        Next c
    Next r

    ' Write the modified formulas back to the range all at once
    sel.Formula = formulas

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub


Sub LockCellReference()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range
    Dim formulas As Variant
    Dim i As Long, j As Long

    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    Set sel = Selection

    ' Read all formulas into an array
    formulas = sel.Formula

    ' Loop through the array and make references absolute
    For i = 1 To UBound(formulas, 1)
        For j = 1 To UBound(formulas, 2)
            If Left(formulas(i, j), 1) = "=" Then
                ' Convert to absolute reference
                formulas(i, j) = Application.ConvertFormula(formulas(i, j), xlA1, xlA1, xlAbsolute)
            End If
        Next j
    Next i

    ' Write back the array all at once
    sel.Formula = formulas

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

Public Sub CycleFormatting()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    Dim sel As Range, firstCell As Range
    Dim nextStyle As Integer
    Dim BLUE_COLOR As Long, RED_COLOR As Long, LIGHTBLUE_COLOR As Long

    ' define colors (RGB can't be used in Const)
    BLUE_COLOR = RGB(0, 32, 96)
    RED_COLOR = RGB(153, 0, 0)
    LIGHTBLUE_COLOR = RGB(220, 228, 244)

    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    Set sel = Selection
    Set firstCell = sel.Cells(1, 1)   ' use first cell to determine next style

    ' Determine next style safely
    If firstCell.Interior.Pattern <> xlNone And firstCell.Interior.Color = BLUE_COLOR Then
        nextStyle = 2   ' move to red font
    ElseIf firstCell.Font.Color = RED_COLOR Then
        nextStyle = 3   ' move to light-blue fill
    ElseIf firstCell.Interior.Pattern <> xlNone And firstCell.Interior.Color = LIGHTBLUE_COLOR Then
        nextStyle = 0   ' reset (clear)
    Else
        nextStyle = 1   ' first style (dark blue fill)
    End If

    ' Apply chosen style to full selection
    With sel
        .Font.Name = "Garamond"
        Select Case nextStyle
            Case 1  ' Dark blue fill, white font, bold
                .Interior.Pattern = xlSolid
                .Interior.Color = BLUE_COLOR
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone

            Case 2  ' Red font, no fill, bold, underlined
                .Interior.Pattern = xlNone
                .Font.Color = RED_COLOR
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone

            Case 3  ' Light blue fill, bold, top & bottom borders
                .Interior.Pattern = xlSolid
                .Interior.Color = LIGHTBLUE_COLOR
                .Font.Color = RGB(0, 0, 0)
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleNone
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With

            Case 0  ' Reset / clear
                .Interior.Pattern = xlNone
                .Font.Color = RGB(0, 0, 0)
                .Font.Bold = False
                .Font.Underline = xlUnderlineStyleNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
        End Select
    End With

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

' Go to the first cell referenced in the current cell's formula (cross-sheet)
Public Sub GoToPreviousReference()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------

    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    If Selection.Cells.Count = 0 Then GoTo CleanExit

    Dim f As String, parts() As String, sheetName As String, addr As String
    Dim Target As Range

    f = Selection.Cells(1, 1).Formula
    If Len(f) = 0 Then GoTo CleanExit
    If Left(f, 1) = "=" Then f = Mid(f, 2)

    On Error Resume Next
    ' Extract first reference (simplest approach: split on !)
    If InStr(f, "!") > 0 Then
        parts = Split(f, "!")
        sheetName = Replace(parts(0), "'", "")
        addr = parts(1)
        Set Target = Worksheets(sheetName).Range(addr)
    Else
        ' Single-sheet reference
        Set Target = Selection.Cells(1, 1).Precedents
    End If
    On Error GoTo CleanExit

    If Not Target Is Nothing Then
        Target.Worksheet.Activate
        Target.Select
    End If

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub


Public Sub GoToNextDependent()
    '---------Optimization-----------
    Dim calcMode As XlCalculation, oldStatus As Boolean
    On Error GoTo CleanExit

    ' Save current settings
    calcMode = Application.Calculation
    oldStatus = Application.DisplayStatusBar

    ' Optimize
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    '---------End_Optimization---------
    
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    
    Dim ws As Worksheet, c As Range, found As Boolean
    Dim currAddress As String
    currAddress = "'" & Selection.Worksheet.Name & "'!" & Selection.Address(False, False)
    
    found = False
    ' Loop through all open sheets and cells with formulas
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        For Each c In ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            If InStr(1, c.Formula, currAddress, vbTextCompare) > 0 Then
                c.Worksheet.Activate
                c.Select
                found = True
                Exit For
            End If
        Next c
        On Error GoTo CleanExit
        If found Then Exit For
    Next ws

' Restore settings
CleanExit:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldStatus
End Sub

