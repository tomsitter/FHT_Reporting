Sub CCO_Report_Summary()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With Sheets("CCO_Report")
        'count the number of rows
        Sheets("CCO_Report").Range("A1:z" & (20000)).Clear
        AllEntries = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row
        
        Sheets("temp").Select
        Sheets("temp").Range("A1:L" & AllEntries).Copy
        Sheets("CCO_Report").Select
        Sheets("CCO_Report").Range("A1").Select
        Sheets("CCO_Report").Paste
        
        'rename/create new headings
        .Range("A1").Value = ""
        .Range("G1").Value = "Breast - Eligible"
        .Range("H1").Value = "Breast - Status"
        .Range("I1").Value = "Cervical - Eligible"
        .Range("J1").Value = "Cervical - Status"
        .Range("K1").Value = "Colorectal - Eligible"
        .Range("L1").Value = "Colorectal - Status"
        .Range("M1").Value = "Reviewed by RN"
        .Range("N1").Value = "Due?"
        .Range("O1").Value = "Receptionist Notes"
        
        'count number of "Action" for each patient
        For i = 2 To AllEntries
            action_count = Application.WorksheetFunction.CountIf(Range(Cells(i, 7), Cells(i, 12)), "Action")
            Debug.Print action_count
            .Range("M" & i).Value = action_count
          
        Next i
        
        'delete those patients with 0 in "review patient" column
        RowNum = 2
        
        For j = 2 To AllEntries
            If .Range("M" & RowNum).Value = 0 Then
                .Range("M" & RowNum).EntireRow.Delete
            Else
                RowNum = RowNum + 1
            End If
        Next j
        
        'create table for actionable patients
        AllActionable = Sheets("CCO_Report").Cells(Rows.Count, 1).End(xlUp).Row
        Dim tbl As ListObject
        Set Rng = .Range(Cells(1, 1), Cells(AllActionable, 15))
    
        Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Rng, xlYes, xlYes)
        tbl.TableStyle = "TableStyleMedium15"
        'change height of rows
        Rows("1:" & AllActionable).RowHeight = 25
        'hide columns
        .Columns("C").Hidden = True
        .Columns("E").Hidden = True
        .Columns("C").Hidden = True
        .Columns("M").Hidden = True
        
        'color "labelled" cells
        Set LabelRng = Range(Cells(1, 7), Cells(AllActionable, 12))
        
        For Each cell In LabelRng
            If cell.Value = "Action" Then cell.Interior.ColorIndex = 3
            If cell.Value = "Normal" Then cell.Interior.ColorIndex = 4
            If cell.Value = "Review" Then cell.Interior.ColorIndex = 6
        Next
    
        'footer
        ActiveSheet.PageSetup.RightFooter = Sheets("Instructions").Range("B3").Value
        
    End With

End Sub

Sub Exclude_Tests()

    Dim pssHINs As Range, ccoHINs As Range, found As Range
    Dim cervical As Range, breast As Range, colorectal As Range
    
    With Sheets("Exclusions")
        Set pssHINs = .Range(.Range("A2"), .Range("A2").End(xlDown))
        Set cervical = .Range(.Range("B2"), .Range("B2").End(xlDown))
        Set breast = .Range(.Range("C2"), .Range("C2").End(xlDown))
        Set colorectal = .Range(.Range("D2"), .Range("D2").End(xlDown))
    End With
    
    With Sheets("CSV")
        Set ccoHINs = .Range(.Range("C5"), .Range("C5").End(xlDown))
    
        For Each cell In pssHINs
            Set found = ccoHINs.Find(cell.Value)
            If Not found Is Nothing Then
                excludeCervical = cervical.Cells(cell.Row - 1)
                If excludeCervical = True Then
                    .Cells(found.Row, 8).Value = "Excluded"
                End If
                
                excludeBreast = breast.Cells(cell.Row - 1)
                If excludeBreast = True Then
                    .Cells(found.Row, 10).Value = "Excluded"
                End If
                excludeColorectal = colorectal.Cells(cell.Row - 1)
                If excludeColorectal = True Then
                    .Cells(found.Row, 12).Value = "Excluded"
                End If
            End If
            
        Next cell
        
    End With
    

End Sub

Sub CSV_Import(worksheetName)
    Dim vFileName
    
    On Error GoTo ErrorHandle
    vFileName = Application.GetOpenFilename("Text Files (*.csv),*.csv")
    
    If vFileName = False Or Right(vFileName, 3) <> "csv" Then
       GoTo BeforeExit
    End If
    
    Application.ScreenUpdating = False
    
    Dim wbkS As Workbook
    Dim wshS As Worksheet
    Dim wshT As Worksheet
    Set wshT = Worksheets(worksheetName)
    
    Set wbkS = Workbooks.Open(Filename:=vFileName)
    Set wshS = wbkS.Worksheets(1)
    wshT.Range("A1:Z1000").Clear
    wshS.UsedRange.Copy Destination:=wshT.Range("A1")
    wbkS.Close SaveChanges:=False
        GoTo Complete
    
Complete:
    MsgBox "Complete: " + vFileName + " has been Imported"
    
BeforeExit:
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandle:
    MsgBox Err.Description
    Resume BeforeExit

End Sub

Sub PSS_Import()
    ' Import the PSS Exclusion code search and perform data transformations
    ' 1) Map PSS HIN to CCO HIN (e.g. ON 1234 567 890 -> 1234567890)
    ' 2) Map Latest Exclusion Code Date to True/False based off CutOff Range

    Call CSV_Import("Exclusions")
    
    Call PSS_HIN_To_CCO
    
    Call MapCutOffs


End Sub

Sub PSS_HIN_To_CCO()
    
    Dim HINs As Range
    
    Set HINs = Sheets("Exclusions").Range("A2:A5000")
    
    For Each cell In HINs
        cell.Value = "'" & Replace(Right(cell.Value, 12), " ", "")
    Next cell
    
End Sub

Sub MapCutOffs()
    Dim cutoff As Range
    Dim cervical As Range, breast As Range, colorectal As Range
    Dim cancers As Range
    
    Set cutoff = Range("cutoff")
    With Sheets("Exclusions")
        ' Set cancers = Sheets("Exclusions").Range("B2:B65565,C2:C65565,D2:D65565")
        Set cervical = .Range(.Range("B2"), .Range("B2").End(xlDown))
        Set breast = .Range(.Range("C2"), .Range("C2").End(xlDown))
        Set colorectal = .Range(.Range("D2"), .Range("D2").End(xlDown))
        
        Call checkCutoff(cervical, cutoff.Value)
        Call checkCutoff(breast, cutoff.Value)
        Call checkCutoff(colorectal, cutoff.Value)
    End With

End Sub

Sub checkCutoff(Rng As Range, cutoff As Long)

    For Each cell In Rng
        If cell.Value <= cutoff Or cell.Value = "never done" Then
            cell.Value = False
        Else
            cell.Value = True
        End If
    Next cell

End Sub
