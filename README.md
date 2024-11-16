# ProtoSheet

ProtoSheet (C)Copyright Stephen Goldsmith 2006-2024. All rights reserved.

Distributed at http://aircraftsystemsafety.com/code/

Eclipse Public License - v 2.0
THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
https://www.eclipse.org/legal/epl-2.0/

Whilst PivotTable is an excellent feature for data analysis, it is not a good solution for generating reports
that need to pull data from various worksheets and merge this into a report template that needs to grow
depending on the amount of data found. With PivotTable you are stuck with filter options that will appear when
printed and will overwrite data below the PivotTable. I therefore needed a simple solution to define a report
prototype and fill that report with matching data ready for printing or copying into a Word document. This is
what ProtoSheet has been created to achieve.

To use, create a prototype worksheet with the content and format you want in the report. Create a column that
will be used to contain commands that this tool will process. It is suggested that you set the background
color of this column (such as to gray) so that it stands apart from the report, and also set a conditional
format such that any row starting with "//" or "#" is shown in a different font color (such as green) so that
comments are clearly distinguished from commands. This command column can either be the first column ("A") and
to the left of the report, or it can be placed to the right of the report. If it is the first column, you must
specify which columns in the prototype worksheet to process so that the script knows how wide the report is.
This can be done either by specifying the start and end column in the arguments to this procedure or by
specifying the number of columns after the command column using the 'COLUMNS' command in a row of the command
column.
