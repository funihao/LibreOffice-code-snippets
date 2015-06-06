'************************************************************************************
'**
'** This LibreOffice / OpenOffice code snippet does the following:
'**
'** 1) Parses sub-directories and files from a base search path
'**
'** 2) Opens all .xls documents found in the base path INVISIBLE, so processing
'**    is faster and the user is not irritated by 200 windows popping up sequentially
'**    when parsing a large number of documents.
'**
'** 3) Gets the content of cell 0,0
'**
'** 4) Displays the text all cell 0,0's of the documents found in a messagebox
'**
'** 6. June 2015, Martin Sauter
'** 
'************************************************************************************

Option Explicit

Sub Main
  Dim s As String                     'Temporary string
  Dim sFileName As String             'Last name returned from DIR
  Dim i As Integer                    'Count number of dirs and files
  Dim sPath                           'Current path with path separator at end
  Dim sBasePath
  Dim sSearchName

  ' variables to to load and access a calc / xls sheet
  Dim CalcDoc 
  Dim Dummy(0) As New com.sun.star.beans.PropertyValue
  Dim Sheet
  Dim Cell
 
  'IMPORTANT: the path string must have a "/" at the end!"
  sBasePath = "/home/martin/x/y/z/"

  'The search name can be omitted or set as below if only looking for documents of a certain type
  'In this case Excel documents are search which are then opened by Libreoffice and converted automatically
  'so the content of the cells can be accessed from the Macro without dealing with the proprietary 
  'data format.
  sSearchName = "*.xls"
  sPath = sBasePath + sSearchName

  'Search all sub-directories 
  '===============================
  sFileName = Dir(sBasePath, 16)      'directory rather than what it contains
  i = 0                               'Initialize the variable
  Do While (sFileName <> "")          'While something returned
    i = i + 1                         'Count the directories
    s = s & "Dir " & CStr(i) &_
        " = " & sFileName & CHR$(10)  'Store in string for later printing
    sFileName = Dir()                 'Get the next directory name
  Loop

  'search all files in the main directory (NOT (!) in the subdirectories)
  '======================================================================
  
  i = 0                               'Start counting over for files
  sFileName = Dir(sPath, 0)           'Get files this time!
  Do While (sFileName <> "")
    i = i + 1
  
    ' Let's try to open that Excel file... 
    Dummy(0).Name="Hidden"
    Dummy(0).Value=TRUE  
    CalcDoc = StarDesktop.loadComponentFromURL(ConvertToURL(sBasePath) + sFileName, "_blank", 0, Dummy())    
    Sheet = CalcDoc.Sheets(0)
    Cell = Sheet.getCellByPosition(0, 0)

    ' Add content of cell 0,0 (=A1) to the output string
    ' that is shown in a messagebox once all files have been searched
    s = s & Cell.String & CHR$(10)
       
    CalcDoc.close(true)
            
    sFileName = Dir()
    
  Loop
  
  
  'Now show the fruits of the serach
  MsgBox s, 0, ConvertToURL(sPath)
End Sub

