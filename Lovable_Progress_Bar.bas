Attribute VB_Name = "Lovable_Progress_Bar"
Sub Lovable_Progress_Bar()
    On Error Resume Next

    ' === DEFAULT CONFIGURATION ===
    Dim startOffset As Integer: startOffset = 1       ' Number of slides to skip at start
    Dim endOffset As Integer: endOffset = 1           ' Number of slides to skip at end
    Dim barHeight As Single: barHeight = 6            ' Height of the progress bar in pixels
    Dim transparency As Single: transparency = 0.3   ' Transparency of the bar (0=solid,1=fully transparent)
    Dim mode As String: mode = "multi"                ' Mode of progress bar: "single" or "multi"
    
    ' COMMON VARIABLES WITH DEFAULTS
    Dim cornerRadius As Single: cornerRadius = 3      ' Radius of rounded corners
    Dim margin As Single: margin = 5                  ' Margin from slide edges
    Dim barColor As Long: barColor = RGB(191, 191, 191) ' Color of the bar
    
    ' MULTI MODE VARIABLES
    Dim gap As Single: gap = 5                        ' Gap between boxes in multi-step mode
    Dim lastBoxDifferentColor As Boolean: lastBoxDifferentColor = False  ' Whether last box has different color
    Dim lastBoxColor As Long: lastBoxColor = RGB(127, 127, 127)          ' Color of last box if different
    
    ' === CONFIRMATION DIALOG FOR DEFAULTS ===
    Dim useDialog As Boolean
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Do you want to use default values?" & vbCrLf & _
                    "Yes ? use defaults" & vbCrLf & _
                    "No ? enter custom values", _
                    vbYesNo + vbQuestion, "Progress Bar Settings")
    If answer = vbNo Then useDialog = True Else useDialog = False
    
    If useDialog Then
        Dim tempStr As String
        Dim cancelled As Boolean: cancelled = False
        
        ' --- COMMON VARIABLES BEFORE MODE ---
        tempStr = InputBox("Enter number of slides to skip at start:", "Start Offset", startOffset)
        If tempStr = "" Then cancelled = True Else startOffset = CInt(tempStr)
        
        If Not cancelled Then
            tempStr = InputBox("Enter number of slides to skip at end:", "End Offset", endOffset)
            If tempStr = "" Then cancelled = True Else endOffset = CInt(tempStr)
        End If
        
        If Not cancelled Then
            tempStr = InputBox("Enter bar color (RGB, e.g., 191,191,191):", "Bar Color", "191,191,191")
            If tempStr <> "" Then
                Dim rgbArr() As String
                rgbArr = Split(tempStr, ",")
                If UBound(rgbArr) = 2 Then
                    barColor = RGB(CInt(rgbArr(0)), CInt(rgbArr(1)), CInt(rgbArr(2)))
                End If
            End If
        End If
        
        If Not cancelled Then
            tempStr = InputBox("Enter progress bar height (pixels):", "Bar Height", barHeight)
            If tempStr = "" Then cancelled = True Else barHeight = CSng(tempStr)
        End If
        
        If Not cancelled Then
            tempStr = InputBox("Enter bar transparency (0=solid, 1=full):", "Transparency", transparency)
            If tempStr = "" Then cancelled = True Else transparency = CSng(tempStr)
        End If
        
        If Not cancelled Then
            tempStr = InputBox("Enter corner radius (pixels):", "Corner Radius", cornerRadius)
            If tempStr = "" Then cancelled = True Else cornerRadius = CSng(tempStr)
        End If
        
        If Not cancelled Then
            tempStr = InputBox("Enter margin from slide edges (pixels):", "Margin", margin)
            If tempStr = "" Then cancelled = True Else margin = CSng(tempStr)
        End If
        
        ' --- ASK FOR MODE USING BUTTONS ---
        If Not cancelled Then
            answer = MsgBox("Choose progress bar mode:" & vbCrLf & _
                            "Yes ? Multi" & vbCrLf & _
                            "No ? Single", _
                            vbYesNo + vbQuestion, "Select Mode")
            If answer = vbYes Then
                mode = "MULTI"
            Else
                mode = "SINGLE"
            End If
        End If
        
        ' --- MULTI MODE VARIABLES ---
        If Not cancelled And UCase(mode) = "MULTI" Then
            tempStr = InputBox("Enter gap between boxes (pixels):", "Gap", gap)
            If tempStr = "" Then cancelled = True Else gap = CSng(tempStr)
            
            If Not cancelled Then
                ' ASK IF LAST BOX SHOULD HAVE DIFFERENT COLOR USING Yes/No
                answer = MsgBox("Do you want the last box to have a different color?" & vbCrLf & _
                                "Yes ? Different color" & vbCrLf & _
                                "No ? Same color as others", _
                                vbYesNo + vbQuestion, "Last Box Color")
                If answer = vbYes Then
                    lastBoxDifferentColor = True
                    ' Ask for last box color
                    tempStr = InputBox("Enter last box color (RGB, e.g., 127,127,127):", "Last Box Color", "127,127,127")
                    If tempStr <> "" Then
                        Dim rgbArr2() As String
                        rgbArr2 = Split(tempStr, ",")
                        If UBound(rgbArr2) = 2 Then
                            lastBoxColor = RGB(CInt(rgbArr2(0)), CInt(rgbArr2(1)), CInt(rgbArr2(2)))
                        End If
                    End If
                Else
                    lastBoxDifferentColor = False
                End If
            End If
        End If
    End If
    
    ' === SETUP PRESENTATION ===
    Dim pres As Presentation
    Set pres = ActivePresentation
    Dim totalSlides As Integer
    totalSlides = pres.Slides.Count
    
    ' --- DELETE ALL PREVIOUS PROGRESS BAR SHAPES ---
    Dim sl As Slide, shpIndex As Integer
    For Each sl In pres.Slides
        For shpIndex = sl.Shapes.Count To 1 Step -1
            If InStr(sl.Shapes(shpIndex).Name, "LPB") > 0 Then
                sl.Shapes(shpIndex).Delete
            End If
        Next shpIndex
    Next sl
    
    ' --- CALCULATE SLIDE RANGE BASED ON OFFSETS ---
    Dim slideStart As Integer, slideEnd As Integer
    slideStart = 1 + startOffset
    slideEnd = totalSlides - endOffset
    If slideStart < 1 Then slideStart = 1
    If slideEnd > totalSlides Then slideEnd = totalSlides
    If slideStart > slideEnd Then Exit Sub
    
    Dim slideRange As Integer
    slideRange = slideEnd - slideStart + 1
    
    ' --- DRAW PROGRESS BAR(S) ---
    Dim X As Integer, j As Integer, s As Shape
    Dim barWidth As Single
    
    If UCase(mode) = "SINGLE" Then
        ' Single continuous bar
        For X = slideStart To slideEnd
            barWidth = (X - slideStart + 1) / slideRange * (pres.PageSetup.SlideWidth - 2 * margin)
            Set s = pres.Slides(X).Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                margin, _
                pres.PageSetup.SlideHeight - barHeight - margin, _
                barWidth, _
                barHeight)
            
            With s
                .Fill.ForeColor.RGB = barColor
                .Fill.transparency = transparency
                .Line.Visible = msoFalse
                .Adjustments.Item(1) = cornerRadius / barHeight
                .Name = "LPB_" & X
            End With
        Next X
        
    ElseIf UCase(mode) = "MULTI" Then
        ' Multi-step bar with optional last box color
        barWidth = ((pres.PageSetup.SlideWidth - 2 * margin) - (slideRange - 1) * gap) / slideRange
        
        For X = slideStart To slideEnd
            For j = slideStart To X
                Set s = pres.Slides(X).Shapes.AddShape( _
                    msoShapeRoundedRectangle, _
                    margin + (j - slideStart) * (barWidth + gap), _
                    pres.PageSetup.SlideHeight - barHeight - margin, _
                    barWidth, _
                    barHeight)
                
                With s
                    If lastBoxDifferentColor And j = X Then
                        .Fill.ForeColor.RGB = lastBoxColor
                    Else
                        .Fill.ForeColor.RGB = barColor
                    End If
                    .Fill.transparency = transparency
                    .Line.Visible = msoFalse
                    .Adjustments.Item(1) = cornerRadius / barHeight
                    .Name = "LPB_" & j
                End With
            Next j
        Next X
    End If
End Sub
