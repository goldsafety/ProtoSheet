Attribute VB_Name = "ProtoSheet"
Option Explicit

'All code is (C)Copyright Stephen Goldsmith 2006-2024. All rights reserved.
'Eclipse Public License - v 2.0
'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
'https://www.eclipse.org/legal/epl-2.0/

Const ErrInArguments = vbObjectError + 1
Const ErrInCommand = vbObjectError + 2

Public Sub ProtoSheet3(PrototypeWorksheetName As String, DestinationWorksheetName As String, CommandColumn As Long, Optional StartColumn As Long = 0, Optional EndColumn As Long = 0, Optional CommentColumn As Long = 0)
    'ProtoSheet (C)Copyright Stephen Goldsmith 2006-2024. All rights reserved.
    'Version 3.3.0 last updated November 2024
    'Distributed at https://github.com/goldsafety/ProtoSheet and http://aircraftsystemsafety.com/code/
    
    'Eclipse Public License - v 2.0
    'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
    'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
    'https://www.eclipse.org/legal/epl-2.0/
    
    'Whilst PivotTable is an excellent feature for data analysis, it is not a good solution for generating reports
    'that need to pull data from various worksheets and merge this into a report template that needs to grow
    'depending on the amount of data found. With PivotTable you are stuck with filter options that will appear when
    'printed and will overwrite data below the PivotTable. I therefore needed a simple solution to define a report
    'prototype and fill that report with matching data ready for printing or copying into a Word document. This is
    'what ProtoSheet has been created to achieve.
    
    'To use, create a prototype worksheet with the content and format you want in the report. Create a column that
    'will be used to contain commands that this tool will process. It is suggested that you set the background
    'color of this column (such as to gray) so that it stands apart from the report, and also set a conditional
    'format such that any row starting with "//" or "#" is shown in a different font color (such as green) so that
    'comments are clearly distinguished from commands. This command column can either be the first column ("A") and
    'to the left of the report, or it can be placed to the right of the report. If it is the first column, you must
    'specify which columns in the prototype worksheet to process so that the script knows how wide the report is.
    'This can be done either by specifying the start and end column in the arguments to this procedure or by
    'specifying the number of columns after the command column using the 'COLUMNS' command in a row of the command
    'column.
    
    'Note: the inline "[n]" format is no longer supported. The new command column format allows drawing data from
    'more than one worksheet and with different filters into the same report. Additionally, the order of arguments
    'for TABLE was changed in 3.2.0 - make sure you update any prototype worksheets accordingly.
    
    'Commands that can be used in the command column:
    'COLUMNS | NumberOfColumns                    Sets how many columns to process in the prototype worksheet
    'END                                          Stops execution (must appear once in the column)
    'FILTER  | Worksheet | [Filter]               Sets a global filter to be used in all following commands
    'TABLE   | Worksheet | [Filter] | [Columns]   Iterates a table adding additional rows for matching data
    
    'The main command that you will use is the TABLE command, which is used as follows:
    '
    '  Builds a table by matching rows in another worksheet and retrieving column data. The data is inserted into
    '  primarily the formula or, if no formula, the value of any cell within the prototype area of that row where
    '  a prototype text entry exists. Additional rows are inserted for each matching row from the source worksheet.
    '  Prototype text entries are specified in braces (curly brackets) like the following:
    '    Text {column1} text {column2}
    '  In formulas, an additional syntax is supported where a modifier can be specified before the opening brace
    '  bracket. Supported modifiers are currently only % which is used to remove surrounding double quote
    '  characters.
    
    'Version History (minor version changes are only shown for the current major version)
    '  1              - Initial development version.
    '  2              - First public release, based on inline hard coded templates.
    '  3.0.0          - Changed to a command column, not compatible with inline format.
    '  3.1.0          - Added COLUMNS command to allow the number of columns to be defined in the prototype.
    '  3.1.1          - Various bug fixes.
    '  3.2.0          - Command column can now be after the prototype rather than only before it. Filter now
    '                   supports wildcards. Templates are now replaced in formulas and not just values. Fixed error
    '                   when no filter was specified for TABLE.
    '  3.3.0 (Stable) - Filter and Columns arguments for TABLE have been swapped and both are now optional. Sub
    '                   and control arguments are now stable and will not be changed without the version number
    '                   increasing to 4.
    
    Dim wsDestination As Worksheet, wsSource As Worksheet, lColFirst As Long, lColLast As Long
    Dim lRowPrototype As Long, lRowTemplate As Long, lRowSource As Long, i As Long, j As Long, k As Long
    Dim sCmd As String, sArgs() As String, sTemplate As String, t As String
    Dim sFilterCol() As String, sFilterVal() As String, sFilterCon() As String, sColumn() As String
    Dim dictValues As Variant, dictFilter As Variant, rCell As Range
    
    'Define the prototype area using the CommandColumn, StartColumn, EndColumn arguments
    If CommandColumn < 1 Or StartColumn < 0 Or EndColumn < 0 Or CommentColumn < 0 Then
        MsgBox "Invalid column number specified for one or more of the column arguments", vbCritical, "ProtoSheet"
        Exit Sub
    ElseIf CommandColumn = CommentColumn Then
        MsgBox "Comment column cannot be the same as the Command column", vbCritical, "ProtoSheet"
        Exit Sub
    ElseIf StartColumn <> 0 Or EndColumn <> 0 Then
        If StartColumn = 0 Or EndColumn = 0 Or StartColumn > EndColumn Then
            MsgBox "Start column has been specified without End column (or vice-versa) or Start column is greater than End column", vbCritical, "ProtoSheet"
            Exit Sub
        ElseIf (CommandColumn >= StartColumn And CommandColumn <= EndColumn) Or (CommentColumn >= StartColumn And CommentColumn <= EndColumn) Then
            MsgBox "Command or Comment column is within the prototype area", vbCritical, "ProtoSheet"
            Exit Sub
        End If
        lColFirst = StartColumn
        lColLast = EndColumn
    ElseIf CommandColumn > 1 Then
        lColFirst = 1
        lColLast = CommandColumn - 1
    Else
        MsgBox "Invalid Command column specified (must be at least 1) or no starting or", vbCritical, "ProtoSheet"
        Exit Sub
    End If
    
    'Make sure there is an 'END' statement in the prototype worksheet
    On Error Resume Next
        Set wsDestination = Worksheets(PrototypeWorksheetName) 'We are just borrowing the wsDestination variable for this check
    On Error GoTo 0
    If wsDestination Is Nothing Then
        MsgBox "The prototype worksheet '" & PrototypeWorksheetName & "' does not exist", vbCritical, "ProtoSheet"
        Exit Sub
    Else
        Set wsDestination = Nothing
    End If
    Set rCell = Worksheets(PrototypeWorksheetName).Columns(CommandColumn).Find("END", , , xlWhole)
    If rCell Is Nothing Then
        MsgBox "The prototype worksheet does not include an 'END' statement in the command column", vbCritical, "ProtoSheet"
        Exit Sub
    End If
    
    'Confirm before replacing the destination worksheet
    On Error Resume Next
        Set wsDestination = Worksheets(DestinationWorksheetName)
    On Error GoTo 0
    If Not wsDestination Is Nothing Then
        i = MsgBox("Are you sure you want to replace the contents of the '" & DestinationWorksheetName & "' worksheet? All current data in this worksheet will be deleted.", vbYesNo Or vbQuestion Or vbDefaultButton2, "ProtoSheet")
        If i <> vbYes Then
            Exit Sub
        End If
        Application.DisplayAlerts = False
            wsDestination.Delete
        Application.DisplayAlerts = True
        Set wsDestination = Nothing
    End If
    
    'Set a global error handler
    Err.Clear
    On Error GoTo Err_ProtoSheet
    
    'Copy the prototype to the destination
    Worksheets(PrototypeWorksheetName).Copy After:=Worksheets(PrototypeWorksheetName)
    ActiveSheet.Name = DestinationWorksheetName
    Set wsDestination = ActiveSheet
    
    'Create a dictionary object for holding the cell references for the row data and for the global table filters
    Set dictValues = CreateObject("Scripting.Dictionary")
    Set dictFilter = CreateObject("Scripting.Dictionary")
    dictValues.CompareMode = vbTextCompare
    dictFilter.CompareMode = vbTextCompare
    
    'Loop over each row of the destination and process the commands in the command column
    lRowPrototype = 0
    lRowTemplate = 0
    Do
        lRowPrototype = lRowPrototype + 1
        lRowTemplate = lRowTemplate + 1
        sCmd = Trim(wsDestination.Cells(lRowTemplate, CommandColumn))
        
        'Strip out comments using either C# or PowerShell/Python/Ruby style (// or #). I have not supported VBA (') style
        'comments as Excel treats this as a special formatting indicator for text. I recommend using the # format as
        'by default Excel will display the menu hotkeys when a user presses '/' unless already editing the cell.
        i = InStrQuoted(1, sCmd, "//")
        If i > 0 Then sCmd = Left(sCmd, i - 1)
        i = InStrQuoted(1, sCmd, "#")
        If i > 0 Then sCmd = Left(sCmd, i - 1)
        
        'Split command into pipe (|) seperated arguments, though ignoring pipe characters in quoted ("") text
        If sCmd <> "" Then
            sArgs = SplitQuoted(sCmd, "|")
            For i = 0 To UBound(sArgs)
                sArgs(i) = Trim(sArgs(i))
            Next
            sArgs(0) = UCase(sArgs(0))
        Else
            ReDim sArgs(0)
            sArgs(0) = ""
        End If
        
        If sArgs(0) = "COLUMNS" Then
            'Sets how many columns to process in the prototype worksheet
            
            'Check there is 1 argument after the command (COLUMNS | NumberOfColumns)
            If UBound(sArgs) <> 1 Then
                Err.Raise ErrInCommand, sCmd, "Incorrect number of arguments for COLUMNS"
            End If
            
            If CommandColumn = 1 Then
                lColLast = sArgs(1) + 1
            Else
                lColLast = sArgs(1)
                If lColLast >= CommandColumn Then
                    Err.Raise ErrInCommand, sCmd, "The number of columns is too high as it would include the command column"
                End If
            End If
        ElseIf sArgs(0) = "FILTER" Then
            'Sets a global filter to be used in all following commands
            
            'Check there is either 1 (FILTER | Worksheet) or 2 (... | Filters) arguments
            If UBound(sArgs) < 1 Or UBound(sArgs) > 2 Then
                Err.Raise ErrInCommand, sCmd, "Incorrect number of arguments for FILTER"
            End If
            
            'First argument is the worksheet name (quote marks are optional as only a single parameter)
            If Left(sArgs(1), 1) = """" And Right(sArgs(1), 1) = """" Then
                sArgs(1) = Mid(sArgs(1), 2, Len(sArgs(1)) - 2)
            End If
            
            If UBound(sArgs) = 1 Then
                If dictFilter.Exists(sArgs(1)) Then
                    dictFilter.Remove sArgs(1)
                Else
                    Err.Raise ErrInCommand, sCmd, "The specified filters cannot be cleared as none have previously been set"
                End If
            ElseIf UBound(sArgs) = 2 Then
                If dictFilter.Exists(sArgs(1)) Then
                    dictFilter.Remove sArgs(1)
                End If
                dictFilter.Add sArgs(1), sArgs(2)
            End If
            
        ElseIf sArgs(0) = "TABLE" Then
            'Search through a worksheet and return values from all matching rows
            
            'Check there is either 1 (TABLE | Worksheet) or 2 (... | Filter) or 3 (... | Columns) arguments
            If UBound(sArgs) < 1 Or UBound(sArgs) > 3 Then
                Err.Raise ErrInCommand, sCmd, "Incorrect number of arguments for TABLE"
            End If
            
            'First argument is the worksheet name (quote marks are optional as only a single parameter)
            'Set a reference to this worksheet in wsSource
            If Left(sArgs(1), 1) = """" And Right(sArgs(1), 1) = """" Then
                sArgs(1) = Mid(sArgs(1), 2, Len(sArgs(1)) - 2)
            End If
            Set wsSource = Worksheets(sArgs(1))
            
            'Second argument (if specified) is a comma-separated list of filters
            'For each, record the filter column ('Col'), condition ('Con') and criteria ('Val')
            If dictFilter.Exists(sArgs(1)) Or UBound(sArgs) >= 2 Then
                If dictFilter.Exists(sArgs(1)) Then
                    If UBound(sArgs) >= 2 Then
                        sFilterCol = SplitQuoted(dictFilter.Item(sArgs(1)) & "," & sArgs(2), ",")
                    Else
                        sFilterCol = SplitQuoted(dictFilter.Item(sArgs(1)), ",")
                    End If
                Else
                    sFilterCol = SplitQuoted(sArgs(2), ",")
                End If
                ReDim sFilterVal(UBound(sFilterCol))
                ReDim sFilterCon(UBound(sFilterCol))
                For i = 0 To UBound(sFilterCol)
                    j = InStrQuoted(1, sFilterCol(i), "=")
                    If j > 0 Then
                        sFilterVal(i) = Trim(Right(sFilterCol(i), Len(sFilterCol(i)) - j))
                        sFilterCol(i) = Trim(Left(sFilterCol(i), j - 1))
                        sFilterCon(i) = "="
                    Else
                        j = InStrQuoted(1, sFilterCol(i), "<>")
                        If j > 0 Then
                            sFilterVal(i) = Trim(Right(sFilterCol(i), Len(sFilterCol(i)) - j - 1))
                            sFilterCol(i) = Trim(Left(sFilterCol(i), j - 1))
                            sFilterCon(i) = "<>"
                        Else
                            Err.Raise ErrInCommand, sCmd, "Invalid or no comparison operator found in '" & sFilterCol(i) & "'"
                        End If
                    End If
                    
                    If Left(sFilterCol(i), 1) = """" And Right(sFilterCol(i), 1) = """" Then
                        sFilterCol(i) = Mid(sFilterCol(i), 2, Len(sFilterCol(i)) - 2)
                        Set rCell = wsSource.Rows(1).Find(sFilterCol(i), , , xlWhole)
                        If rCell Is Nothing Then
                            Err.Raise ErrInCommand, sCmd, "Could not find column '" & sFilterCol(i) & "' in worksheet '" & wsSource.Name & "'"
                        End If
                        sFilterCol(i) = Mid(rCell.Address, 2, InStr(2, rCell.Address, "$") - 2)
                    Else
                        'Check that sFilterCol(i) is a valid column reference
                        On Error GoTo 0
                        Err.Clear
                        On Error Resume Next
                            t = wsSource.Range(sFilterCol(i) & "1")
                        If Err.Number <> 0 Then
                            On Error GoTo Err_ProtoSheet
                            Err.Raise ErrInCommand, sCmd, "Could not read from column " & sFilterCol(i) & " from worksheet '" & wsSource.Name & "'. The most likely source of this error is that a column title has been specified without enclosing in double-quotes so has been interpreted as a column reference. Add double quotes around '" & sFilterCol(i) & "' to search for the column title instead."
                        End If
                        On Error GoTo Err_ProtoSheet
                    End If
                    
                    If Left(sFilterVal(i), 1) = """" And Right(sFilterVal(i), 1) = """" Then
                        sFilterVal(i) = Mid(sFilterVal(i), 2, Len(sFilterVal(i)) - 2)
                    Else
                        Err.Raise ErrInCommand, sCmd, "Missing or unmatched double quotes surrounding '" & sFilterVal(i) & "'"
                    End If
                Next
            End If
            
            'Third argument (if specified) is a comma-separated list of column references or names in quotes
            'For each get a pair of column name and reference
            dictValues.RemoveAll
            If UBound(sArgs) = 3 Then
                sColumn = SplitQuoted(sArgs(3), ",")
                For i = 0 To UBound(sColumn)
                    sColumn(i) = Trim(sColumn(i))
                    
                    'Alternative to implement "Column" AS "Name" format
                    'If Left(sColumn(i), 1) = """" Then
                    '    j = InStr(sColumn(i), """")
                    '    sColumn(i) = Mid(sColumn(i), 2, j - 2)
                    
                    If Left(sColumn(i), 1) = """" And Right(sColumn(i), 1) = """" Then
                        sColumn(i) = Mid(sColumn(i), 2, Len(sColumn(i)) - 2)
                        Set rCell = wsSource.Rows(1).Find(sColumn(i), , , xlWhole)
                        If rCell Is Nothing Then
                            Err.Raise ErrInCommand, sCmd, "Could not find column '" & sColumn(i) & "' in worksheet '" & wsSource.Name & "'"
                        End If
                        dictValues.Add sColumn(i), Mid(rCell.Address, 2, InStr(2, rCell.Address, "$") - 2)
                    Else
                        dictValues.Add sColumn(i), sColumn(i)
                    End If
                Next
            End If
            
            'Iterate through each row of wsSource worksheet to see if it matches the filter
            lRowSource = 2 'Ignore header row
            Do While Application.WorksheetFunction.CountA(wsSource.Rows(lRowSource)) > 0
                'See if this row matches ALL of the filter criteria
                j = 1 'j is the match result, default to TRUE (1)
                For i = 0 To UBound(sFilterCol)
                    On Error GoTo 0
                    Err.Clear
                    On Error Resume Next
                        t = wsSource.Range(sFilterCol(i) & lRowSource)
                    If Err.Number <> 0 Then
                        On Error GoTo Err_ProtoSheet
                        Err.Raise ErrInCommand, sCmd, "Could not read the value of cell " & sFilterCol(i) & lRowSource & " from worksheet '" & wsSource.Name & "'. The most likely source of this error is that a column title has been specified without enclosing in double-quotes so has been interpreted as a column reference. Add double quotes around '" & sFilterCol(i) & "' to search for the column title instead."
                    End If
                    On Error GoTo Err_ProtoSheet
                    
                    If sFilterCon(i) = "=" Then 'And t <> sFilterVal(i)
                        
                        If Left(sFilterVal(i), 1) = "*" And Right(sFilterVal(i), 1) = "*" Then
                            If Len(sFilterVal(i)) < 3 Then
                                'Filter is either "**" or "*" and therefore matches anything
                            ElseIf InStr(t, Mid(sFilterVal(i), 2, Len(sFilterVal(i)) - 2)) = 0 Then
                                j = 0
                            End If
                        ElseIf Left(sFilterVal(i), 1) = "*" Then
                            If Right(t, Len(sFilterVal(i)) - 1) <> Right(sFilterVal(i), Len(sFilterVal(i)) - 1) Then
                                j = 0
                            End If
                        ElseIf Right(sFilterVal(i), 1) = "*" Then
                            If Left(t, Len(sFilterVal(i)) - 1) <> Left(sFilterVal(i), Len(sFilterVal(i)) - 1) Then
                                j = 0
                            End If
                        Else
                            If t <> sFilterVal(i) Then
                                j = 0
                            End If
                        End If
                        
                    ElseIf sFilterCon(i) = "<>" Then 'And t = sFilterVal(i)
                        
                        If Left(sFilterVal(i), 1) = "*" And Right(sFilterVal(i), 1) = "*" Then
                            If Len(sFilterVal(i)) < 3 Then
                                'Filter is either "**" or "*" and therefore matches anything
                            ElseIf InStr(t, Mid(sFilterVal(i), 2, Len(sFilterVal(i)) - 2)) > 0 Then
                                j = 0
                            End If
                        ElseIf Left(sFilterVal(i), 1) = "*" Then
                            If Right(t, Len(sFilterVal(i)) - 1) = Right(sFilterVal(i), Len(sFilterVal(i)) - 1) Then
                                j = 0
                            End If
                        ElseIf Right(sFilterVal(i), 1) = "*" Then
                            If Left(t, Len(sFilterVal(i)) - 1) = Left(sFilterVal(i), Len(sFilterVal(i)) - 1) Then
                                j = 0
                            End If
                        Else
                            If t = sFilterVal(i) Then
                                j = 0
                            End If
                        End If
                        
                    End If
                Next
                If j = 1 Then
                    wsDestination.Rows(lRowTemplate).Insert xlShiftDown, xlFormatFromRightOrBelow
                    wsDestination.Rows(lRowTemplate + 1).Copy wsDestination.Rows(lRowTemplate)
                    
                    'Iterate across each column and fill in from the template
                    For i = lColFirst To lColLast
                        If Application.WorksheetFunction.IsFormula(wsDestination.Cells(lRowTemplate, i)) = True Then
                            t = wsDestination.Cells(lRowTemplate, i).Formula
                        Else
                            t = wsDestination.Cells(lRowTemplate, i)
                        End If
                        
                        Do
                            j = InStr(t, "{")
                            If j > 0 Then
                                k = InStr(j, t, "}")
                                If k = 0 Then
                                    Err.Raise ErrInCommand, sCmd, "Unmatched brace in '" & t & "'"
                                End If
                                sTemplate = Trim(Mid(t, j + 1, k - j - 1))
                                
                                'If no column names were specified in the TABLE command, try and match it here
                                If UBound(sArgs) < 3 Then 'Previously we included "And lRowSource = 2" but we cannot do that as the filter might not match the first row
                                    If dictValues.Exists(sTemplate) = False Then
                                        Set rCell = wsSource.Rows(1).Find(sTemplate, , , xlWhole)
                                        If rCell Is Nothing Then
                                            Err.Raise ErrInCommand, sCmd, "Could not find column '" & sTemplate & "' in worksheet '" & wsSource.Name & "'"
                                        End If
                                        dictValues.Add sTemplate, Mid(rCell.Address, 2, InStr(2, rCell.Address, "$") - 2)
                                    End If
                                Else
                                    If dictValues.Exists(sTemplate) = False Then
                                        Err.Raise ErrInCommand, sCmd, "Column '" & sTemplate & "' has not been specified in the TABLE command columns argument"
                                    End If
                                End If
                                
                                'Check for a modifier before the opening bracket (currently only supports the % modifier)
                                If j > 1 And Application.WorksheetFunction.IsFormula(wsDestination.Cells(lRowTemplate, i)) = True Then
                                    If Mid(t, j - 1, 1) = "%" Then
                                        'The % modifier must immediately preceed the opening brace and be surrounded by double quote characters which will be removed as well as replacing the template
                                        If j < 3 Or k > (Len(t) - 1) Then
                                            Err.Raise ErrInCommand, sCmd, "% modifier used in '" & t & "' but quotes do not immediately surround the {}"
                                        ElseIf Mid(t, j - 2, 1) <> """" Or Mid(t, k + 1, 1) <> """" Then
                                            Err.Raise ErrInCommand, sCmd, "% modifier used in '" & t & "' but quotes do not immediately surround the {}"
                                        End If
                                        t = Left(t, j - 3) & wsSource.Range(dictValues.Item(sTemplate) & lRowSource) & Right(t, Len(t) - k - 1)
                                    Else
                                        t = Left(t, j - 1) & wsSource.Range(dictValues.Item(sTemplate) & lRowSource) & Right(t, Len(t) - k)
                                    End If
                                Else
                                    t = Left(t, j - 1) & wsSource.Range(dictValues.Item(sTemplate) & lRowSource) & Right(t, Len(t) - k)
                                End If
                            End If
                        Loop Until j = 0
                        
                        If Application.WorksheetFunction.IsFormula(wsDestination.Cells(lRowTemplate, i)) = True Then
                            wsDestination.Cells(lRowTemplate, i).Formula = t
                        Else
                            wsDestination.Cells(lRowTemplate, i) = t
                        End If
                    Next
                    
                    lRowTemplate = lRowTemplate + 1
                End If
                
                lRowSource = lRowSource + 1
            Loop
            wsDestination.Rows(lRowTemplate).Delete
            lRowTemplate = lRowTemplate - 1
            
        ElseIf sArgs(0) = "" Or sArgs(0) = "END" Then
            'This will simply cause the loop to pass to the next line or exit
        Else
            Err.Raise ErrInCommand, sCmd, "Invalid command"
        End If
    Loop Until sArgs(0) = "END" Or lRowTemplate = 32767
    
    'Restore the default xlPart rather than xlWhole setting as this is globally set in the main find dialog box
    Set rCell = Worksheets(PrototypeWorksheetName).Columns(1).Find("END", , , xlPart)
    
    i = MsgBox("The worksheet '" & DestinationWorksheetName & "' has been successfully created. Do you want to delete the command " & IIf(CommentColumn > 0, "and comment columns?", "column?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "ProtoSheet")
    If i = vbYes Then
        'Delete the command and comments column from the destination worksheet
        If CommentColumn > 0 And CommentColumn < CommandColumn Then
            wsDestination.Columns(CommandColumn).Delete
            wsDestination.Columns(CommentColumn).Delete
        ElseIf CommentColumn > 0 And CommentColumn > CommandColumn Then
            wsDestination.Columns(CommentColumn).Delete
            wsDestination.Columns(CommandColumn).Delete
        Else
            wsDestination.Columns(CommandColumn).Delete
        End If
    End If
    
    Exit Sub
    
Err_ProtoSheet:
    If Err.Number = ErrInArguments Then
        MsgBox Err.Description, vbCritical, "ProtoSheet"
    ElseIf Err.Number = ErrInCommand Then
        Worksheets(PrototypeWorksheetName).Activate
        Worksheets(PrototypeWorksheetName).Cells(lRowPrototype, CommandColumn).Select
        MsgBox "Error processing: " & sCmd & vbCrLf & Err.Description, vbCritical, "ProtoSheet"
    Else
        MsgBox "An unhandled error occurred (" & Err.Description & ")", vbCritical, "ProtoSheet"
    End If
    On Error GoTo 0
End Sub


Private Function InStrQuoted(Optional start As Long = 1, Optional string1 As String, Optional string2 As String, Optional compare As Variant = Empty, Optional quotechar As String = """") As Long
    'InStrQuoted (C)Copyright Stephen Goldsmith 2024. All rights reserved.
    'Version 1.0.0 last updated November 2024
    'Distributed at https://github.com/goldsafety/ProtoSheet and http://aircraftsystemsafety.com/code/
    
    'Eclipse Public License - v 2.0
    'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
    'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
    'https://www.eclipse.org/legal/epl-2.0/
    
    'This function is equivalent to the built-in InStr() command, however it ignores any matches that resides
    'inside quoted text. The quote character used can be specified, and can be escaped inside a quote by using
    'two quote characters together as is the norm in VBA e.g. "This text is ""quoted"" inside the string".
    
    Dim lngPos As Long, lngQuoteStart As Long, lngQuoteMid As Long, lngQuoteEnd As Long, lngQuotePos As Long, lngFind As Long
    
    If string2 = quotechar Then
        InStrQuoted = InStr(start, string1, string2, compare)
        Exit Function
    ElseIf InStr(1, string2, quotechar, compare) > 0 Then
        'Continue
    End If
    
    lngPos = start
    Do
        lngQuoteStart = InStr(lngPos, string1, quotechar, compare)
        lngFind = InStr(lngPos, string1, string2, compare)
        
        If lngQuoteStart > 0 And lngQuoteStart < lngFind Then
            lngQuotePos = lngQuoteStart + Len(quotechar)
            Do
                lngQuoteMid = InStr(lngQuotePos, string1, quotechar & quotechar, compare)
                lngQuoteEnd = InStr(lngQuotePos, string1, quotechar, compare)
                If lngQuoteMid = lngQuoteEnd Then
                    lngQuotePos = lngQuoteMid + Len(quotechar) * 2
                End If
            Loop Until lngQuoteMid <> lngQuoteEnd Or lngQuoteEnd = 0
            If lngQuoteEnd = 0 Then
                Err.Raise -1
            Else
                lngPos = lngQuoteEnd + Len(quotechar)
            End If
        End If
    Loop Until lngFind = 0 Or lngQuoteStart = 0 Or lngFind < lngQuoteStart
    
    InStrQuoted = lngFind
End Function

Private Function SplitQuoted(expression As String, Optional delimiter As String = " ", Optional limit As Long = -1, Optional compare As Variant = Empty, Optional quotechar As String = """", Optional trimsubstrings As Boolean = False) As Variant
    'SplitQuoted (C)Copyright Stephen Goldsmith 2024. All rights reserved.
    'Version 1.0.0 last updated November 2024
    'Distributed at https://github.com/goldsafety/ProtoSheet and http://aircraftsystemsafety.com/code/
    
    'Eclipse Public License - v 2.0
    'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
    'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
    'https://www.eclipse.org/legal/epl-2.0/
    
    'This function is equivalent to the built-in Split() command, however it ignores any delimeter that resides
    'inside quoted text. The quote character used can be specified, and can be escaped inside a quote by using
    'two quote characters together as is the norm in VBA e.g. "This text is ""quoted"" inside the string".
    
    Dim arr() As String, lngSubstrings As Long, lngPos As Long
    Dim lngQuoteStart As Long, lngQuoteMid As Long, lngQuoteEnd As Long, lngQuotePos As Long, lngDelim As Long
    Dim str As String
    
    lngPos = 1
    lngSubstrings = 0
    Do
        If limit = -1 Or lngSubstrings < limit Then
            lngQuoteStart = InStr(lngPos, expression, quotechar, compare)
            lngDelim = InStr(lngPos, expression, delimiter, compare)
            If lngDelim = 0 Then lngDelim = Len(expression) + 1
        Else
            lngQuoteStart = 0
            lngDelim = Len(expression) + 1
        End If
        
        Do While lngQuoteStart > 0 And lngQuoteStart < lngDelim
            lngQuotePos = lngQuoteStart + Len(quotechar)
            Do
                lngQuoteMid = InStr(lngQuotePos, expression, quotechar & quotechar, compare)
                lngQuoteEnd = InStr(lngQuotePos, expression, quotechar, compare)
                If lngQuoteMid = lngQuoteEnd Then
                    lngQuotePos = lngQuoteMid + Len(quotechar) * 2
                End If
            Loop Until lngQuoteMid <> lngQuoteEnd Or lngQuoteEnd = 0
            Debug.Assert lngQuoteEnd <> 0
            lngQuoteStart = InStr(lngQuoteEnd + Len(quotechar), expression, quotechar, compare)
            lngDelim = InStr(lngQuoteEnd + Len(quotechar), expression, delimiter, compare)
            If lngDelim = 0 Then lngDelim = Len(expression) + 1
        Loop
        
        lngSubstrings = lngSubstrings + 1
        ReDim Preserve arr(lngSubstrings - 1)
        str = Mid(expression, lngPos, lngDelim - lngPos)
        If trimsubstrings Then str = Trim(str)
        arr(lngSubstrings - 1) = str
        lngPos = lngDelim + Len(delimiter)
    Loop Until lngPos > Len(expression)
    
    SplitQuoted = arr
End Function

'Below are test functions used during the development of ProtoSheet
'All code is (C)Copyright Stephen Goldsmith 2024. All rights reserved.
'Eclipse Public License - v 2.0
'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
'https://www.eclipse.org/legal/epl-2.0/

Private Sub PrintErrorCodes()
    Dim l As Long
    For l = 1 To 512
        If Error(l) <> Error(1) Then
            Debug.Print l & ": " & Error(l)
        End If
    Next
End Sub

Private Sub TestSplitQ()
    Dim str As String
    Dim arr() As String
    Dim i As Long
    
    Debug.Assert InStrQuoted(, "ab ""cd, ef"" gh ""ij, kl"", mn", ",") = 24
    
    str = """One, Two"", ""sneaky"", Three"
    
    arr = SplitQuoted(str, ",")
    For i = 0 To UBound(arr)
        Debug.Print i & ": """ & arr(i) & """"
    Next
End Sub
