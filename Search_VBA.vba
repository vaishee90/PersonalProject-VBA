'References:
'https://stackoverflow.com/questions/31414106/get-list-of-excel-files-in-a-folder-using-vba
'https://msdn.microsoft.com/en-us/library/office/gg549168(v=office.14).aspx

Sub ConsolidatedWorkbook()

    Dim ConsolidatedSheet As Worksheet
    Dim FolderPath As String
    Dim SelectedFiles() As Variant
    Dim NewRow As Long
    Dim FileName As Variant
    Dim WorkBk As Workbook
    Dim ColnWorkBk As Workbook
    Dim CompWorkbk As Workbook
    Dim SourceRange As Range
    Dim Category As String
    Dim CompFileName As String
    Dim CompRange As Range
    Dim DestRange As Range
    Dim CatDestRange As Range
    Dim i As Integer
    Dim compVal As String
    Dim compCount As Long
    
    Dim StartTime As Double
    Dim MinutesElapsed As String

    Dim myFile       As Object
    Dim myFSO        As Object
    Dim myFolder     As Object
    Dim myFiles      As Object

    'Creating a new worksheet
    Set ConsolidatedSheet = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    'Please change the folder path to your local path
    FolderPath = 'folder path to be filled
    NewRow = 1

    'using FileSystemObject to get the files from a folder
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Set myFolder = myFSO.GetFolder(FolderPath)
    Set myFiles = myFolder.Files
 
    'to put the filenames from the folder and put it into an array 
    ReDim nameArray(1 To myFiles.Count)

    i = 1

    For Each myFile In myFiles
        nameArray(i) = myFile.Name
        i = i + 1
    Next   

    'I used the file (July_2017) to fetch the column names.
    'SourceRange - range to be copied from the source files
    'DestRange - range in the destination file for pasting the content from source

    Set ColnWorkBk = Workbooks.Open(FolderPath & Dir(FolderPath & "July_2017.csv"))
    Set SourceRange = ColnWorkBk.Worksheets(1).Range("C1:BW1")
    Set DestRange = ConsolidatedSheet.Range("A" & NewRow)
    Set DestRange = DestRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
    DestRange.Value = SourceRange.Value

    'incrementing new row value after having pasted the contents
    NewRow = NewRow + DestRange.Rows.Count
    ColnWorkBk.Close savechanges:=False

    'traverse through each filename in the array
    For Each FileName In nameArray
        'only traverse filenames following a certain pattern
        If FileName Like "FA*.csv" Then

            Set WorkBk = Workbooks.Open(FolderPath & FileName)
            Set SourceRange = WorkBk.Worksheets(1).Range("A1:BQ" & Cells(Rows.Count, 1).End(xlUp).Row)
            Set DestRange = ConsolidatedSheet.Range("B" & NewRow)
            Set CatDestRange = ConsolidatedSheet.Range("A" & NewRow)
            Set CatDestRange = CatDestRange.Resize(SourceRange.Rows.Count, 1)
            Set DestRange = DestRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            DestRange.Value = SourceRange.Value
            WorkBk.Close

            'comparing filename against another excel sheet to fetch values for a 
            'particular column in the destination file with respect to the filename
            Set CompWorkbk = Workbooks.Open(FolderPath & Dir(FolderPath & "Fields to check.xlsx"))
            Set CompRange = CompWorkbk.Worksheets(1).Range("A2:B" & Cells(Rows.Count, 1).End(xlUp).Row)
            compCount = CompRange.Rows.Count

            For i = 2 To compCount
                compVal = Range("A" & i).Value
                If compVal = FileName Then
                    Category = Range("B" & i).Value
                    Exit For
                End If
            Next i

            CompWorkbk.Close
            CatDestRange.Value = Category
            NewRow = NewRow + DestRange.Rows.Count
        End If
   Next FileName

End Sub
