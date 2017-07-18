Imports System.Configuration
Imports Smartsheet.Api
Imports Smartsheet.Api.Models

Module rwsheet

    Class Program

        ' The API identifies columns by Id, but it's more convenient to refer to column names
        Shared columnMap As New Dictionary(Of String, Long)()     ' Map from friendly column name to column Id 

        Public Shared Sub Main(args As String())

            ' Get API access token from App.config file or environment
            Dim accessToken As String = ConfigurationManager.AppSettings("AccessToken")
            If String.IsNullOrEmpty(accessToken) Then
                accessToken = Environment.GetEnvironmentVariable("SMARTSHEET_ACCESS_TOKEN")
            End If
            If String.IsNullOrEmpty(accessToken) Then
                Throw New Exception("Must set API access token in App.conf file or 'SMARTSHEET_ACCESS_TOKEN' environment variable")
            End If

            ' Initialize client
            Dim ss As SmartsheetClient = New SmartsheetBuilder().SetAccessToken(accessToken).Build()

            ' Get sheet Id from App.config file 
            Dim sheetIdString As String = ConfigurationManager.AppSettings("SheetId")
            Dim sheetId As Long = Long.Parse(sheetIdString)

            ' Load the entire sheet
            Dim sheet As Sheet = ss.SheetResources.GetSheet(sheetId, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            Console.WriteLine("Loaded " & sheet.Rows.Count & " rows from sheet: " & sheet.Name)

            ' Build column map for later reference
            For Each column As Column In sheet.Columns
                columnMap.Add(column.Title, column.Id)
            Next

            ' Accumulate rows needing update here
            Dim rowsToUpdate As New List(Of Row)()

            For Each row As Row In sheet.Rows
                Dim rowToUpdate As Row = evaluateRowAndBuildUpdates(row)
                If rowToUpdate IsNot Nothing Then
                    rowsToUpdate.Add(rowToUpdate)
                End If
            Next

            ' Finally, write all updated cells back to Smartsheet 
            Console.WriteLine("Writing " & rowsToUpdate.Count & " rows back to sheet id " & sheet.Id)
            ss.SheetResources.RowResources.UpdateRows(sheetId, rowsToUpdate)

            Console.WriteLine("Done (Hit enter)")
            Console.ReadLine()
        End Sub


        ' Replace the body of this loop with your code
        ' This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero 
        ' Return a new Row with updated cell values, else null to leave unchanged
        Private Shared Function evaluateRowAndBuildUpdates(sourceRow As Row) As Row
            Dim rowToUpdate As Row = Nothing

            ' Find cell we want to examine
            Dim statusCell As Cell = getCellByColumnName(sourceRow, "Status")

            If statusCell.DisplayValue = "Complete" Then
                Dim remainingCell As Cell = getCellByColumnName(sourceRow, "Remaining")

                If remainingCell.DisplayValue <> "0" Then                         ' Skip if "Remaining" is already zero
                    Console.WriteLine("Need to update row # " & sourceRow.RowNumber)
                    ' We need to update this row, so create list of cells to update and containing row

                    ' Build new cell
                    Dim cellToUpdate = New Cell()
                    cellToUpdate.ColumnId = columnMap("Remaining")
                    cellToUpdate.Value = 0

                    ' Build list of cells to update
                    Dim cellsToUpdate = New List(Of Cell)()
                    cellsToUpdate.Add(cellToUpdate)

                    ' Build containing row
                    rowToUpdate = New Row()
                    rowToUpdate.Id = sourceRow.Id
                    rowToUpdate.Cells = cellsToUpdate
                End If
            End If
            Return rowToUpdate
        End Function


        ' Helper function to find cell in a row
        Private Shared Function getCellByColumnName(row As Row, columnName As String) As Cell
            Return row.Cells.First(Function(cell) cell.ColumnId = columnMap(columnName))
        End Function
    End Class

End Module
