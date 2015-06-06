'************************************************************************************
'**
'** This LibreOffice / OpenOffice code snippet creates a new writer document and then 
'** creates a number of paragraph styles with interesting features such as:
'**
'** * do not split over two pages
'** 
'** * Keep thais paragraph together with the next paragraph. If that one wraps 
'**   to the next page, this one has to wrap, too!
'** 
'** * Ident, center, bold
'**
'** For details on paragraph properties that can be set have a look at:
'** https://www.openoffice.org/api/docs/common/ref/com/sun/star/style/ParagraphProperties.html
'**
'** Some of this code was taken "OpenOffice.org Macros Explained" 3.0 by Andrew Pitonyak
'** http://www.pitonyak.org/oo.php
'** 
'** 6. June 2015 Martin Sauter
'**
'************************************************************************************

Option Explicit

Const my_DefaultFontName = "Times New Roman"

Sub Main

  Dim Url
  Dim Doc
  Dim Dummy()
  Dim TextCursor
  
  DIM NewParagraph
  NewParagraph = com.sun.star.text.ControlCharacter.APPEND_PARAGRAPH
    
  ' Open writer document and output calc content
  Url = "private:factory/swriter"
  Doc = StarDesktop.loadComponentFromURL(Url, "_blank", 0, Dummy())
  
  'Create the paragraph styles we need in the new document
  my_CreateParStylesForMyDoc(Doc)
  
  my_SetPageSizeA4(Doc)

  'Get a text cursor for inserting stuff at the end of the document
  TextCursor = Doc.getText.getEnd()

  'Now let's output some text with some of the new paragraph styles
  '=================================================================
  
  TextCursor.ParaStyleName = "my-NameStyle"
  Doc.getText.insertString(TextCursor, "Some Banner Text", False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)
  
  TextCursor.ParaStyleName = "my-TextNoIdentBold"
  Doc.getText.insertString(TextCursor, "Some Text: ", False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)
        
  TextCursor.ParaStyleName = "my-Text1Ident"
  Doc.getText.insertString(TextCursor, "More Text", False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)

  'Some nice formatting for the end of something
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)       
  TextCursor.ParaStyleName = "my-CenterBold"
  Doc.getText.insertString(TextCursor, "____________", False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)
  TextCursor.ParaStyleName = "my-TextNoIdentBold"
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)
  Doc.getText.InsertControlCharacter(TextCursor, NewParagraph, False)

End Sub


'*********************************************************************************
'**
'** Create all paragraph styles for the writer document
'**
'*********************************************************************************
Sub my_CreateParStylesForMyDoc(Doc)
  REM Tab stops are set in the paragraph style
  ' 1/4 of an inch
  DIM tabStopLoc%
  DIM oProps as Object
  
  tabStopLoc% = 2540 / 4

  'For details on paragraph properties that can be set have a look at:
  'https://www.openoffice.org/api/docs/common/ref/com/sun/star/style/ParagraphProperties.html

  'Text no iddent
  oProps = Array (my_CreateProperty("ParaLeftMargin", CLng(2540 * 0.0)), _
     my_CreateProperty("CharHeight", 12), _
     my_CreateProperty("CharFontName",my_DefaultFontName), _
     my_CreateProperty("CharWeight", com.sun.star.awt.FontWeight.NORMAL) )
  my_CreateParStyle(Doc, "my-TextNoIdent", oProps())

  'Name/Header Paragraph Style
  oProps = Array (my_CreateProperty("ParaLeftMargin", CLng(2540 * 0.0)), _
     my_CreateProperty("CharHeight", 14), _
     my_CreateProperty("CharFontName",my_DefaultFontName), _
     my_CreateProperty("CharWeight", com.sun.star.awt.FontWeight.BOLD) )
  my_CreateParStyle(Doc, "my-NameStyle", oProps())
  
  'Text no iddent and bold
  '*****************************************************************
  'Important Property: Keep this paragraph together with the next 
  'paragraph. If that one wraps to the next page, this one has to 
  'wrap, too!
  '*****************************************************************
  oProps = Array (my_CreateProperty("ParaLeftMargin", CLng(2540 * 0.0)), _
     my_CreateProperty("CharHeight", 12), _
     my_CreateProperty("CharFontName",my_DefaultFontName), _
     my_CreateProperty("ParaKeepTogether", TRUE), _
     my_CreateProperty("CharWeight", com.sun.star.awt.FontWeight.BOLD) )
  my_CreateParStyle(Doc, "my-TextNoIdentBold", oProps())

  'Text 1st ident
  '*****************************************************************
  'IMPORTANT Property: Do not split "ParaSplit" over several pages!
  'Instead, move paragraph to a new page and leave some of the 
  'previous page unused!
  '*****************************************************************
  oProps = Array (my_CreateProperty("ParaLeftMargin", CLng(2540 * 0.2)), _
     my_CreateProperty("CharHeight", 12), _
     my_CreateProperty("CharFontName",my_DefaultFontName), _
     my_CreateProperty("ParaSplit", FALSE), _
     my_CreateProperty("ParaAdjust", com.sun.star.style.ParagraphAdjust.BLOCK), _
     my_CreateProperty("CharWeight", com.sun.star.awt.FontWeight.NORMAL), _
     my_CreateProperty("ParaBottomMargin", 200))
  my_CreateParStyle(Doc, "my-Text1Ident", oProps())

  'Center bold for the '------------' 
  oProps = Array (my_CreateProperty("ParaLeftMargin", CLng(2540 * 0.0)), _
     my_CreateProperty("CharHeight", 33), _
     my_CreateProperty("CharFontName",my_DefaultFontName), _
     my_createProperty("ParaAdjust", 3), _
     my_CreateProperty("CharWeight", com.sun.star.awt.FontWeight.BOLD) )
  my_CreateParStyle(Doc, "my-CenterBold", oProps())
 
End sub


'*********************************************************************************
'**
'** Create and return a PropertyValue structure.
'**
'*********************************************************************************
Function my_CreateProperty( Optional cName As String, Optional uValue ) As com.sun.star.beans.PropertyValue
   Dim oPropertyValue As New com.sun.star.beans.PropertyValue
   If Not IsMissing( cName ) Then
      oPropertyValue.Name = cName
   EndIf
   If Not IsMissing( uValue ) Then
      oPropertyValue.Value = uValue
   EndIf
   my_CreateProperty() = oPropertyValue
End Function 


'*********************************************************************************
'**
'** This function creates a new paragraph style based on the 
'** Properties given to it in the oProps() object array.
'** Individual entries in the object array are created with the
'** my_CreateProperty() function which is in turn called
'** from my_CreateParStylesForMyDoc
'**
'*********************************************************************************
Sub my_CreateParStyle(Doc, sStyleName$, oProps())
  Dim i%, j%
  Dim oFamilies
  Dim oStyle
  Dim oStyles
  Dim tabStops%
  
  oFamilies = Doc.StyleFamilies
  oStyles = oFamilies.getByName("ParagraphStyles")
  If oStyles.HasByName(sStyleName) Then
    Exit Sub
  End If
  oStyle = Doc.createInstance("com.sun.star.style.ParagraphStyle")
  For i=LBound(oProps) To UBound(oProps)
    If oProps(i).Name = "ParentStyle" Then
      If oStyles.HasByName(oProps(i).Value) Then
        oStyle.ParentStyle = oProps(i).Value
      Else
        Print "Parent paragraph style (" & oProps(i).Value & _
              ") does not exist, ignoring parent"
      End If
    ElseIf oProps(i).Name = "ParaTabStops" Then
      tabStops = oProps(i).Value
      Dim tab(0 To 19) As New com.sun.star.style.TabStop
      For j =LBound(tab) To UBound(tab)
        tab(j).Alignment = com.sun.star.style.TabAlign.LEFT
        tab(j).DecimalChar = ASC(".")
        tab(j).FillChar = 32
        tab(j).Position = (j+1) * tabStops
      Next
      oStyle.ParaTabStops = tab
    ElseIf oProps(i).Name = "FollowStyle" Then
      If oStyles.HasByName(oProps(i).Value) OR oProps(i).Value = sStyleName Then
        oStyle.setPropertyValue(oProps(i).Name, oProps(i).Value)
      Else
        Print "Next paragraph style (" & oProps(i).Value & _
              ") does not exist, ignoring for style " & sStyleName
      End If
    Else
      oStyle.setPropertyValue(oProps(i).Name, oProps(i).Value)
    End If
  Next
  oStyles.insertByName(sStyleName, oStyle)
End Sub


'****************************************************************
'*
'*
'****************************************************************
Sub my_SetPageSizeA4(oDoc)
    
    Dim oStyle
    
    oStyle = oDoc.StyleFamilies.getByName("PageStyles").getByName("Default Style")
	' units of 1/1000 cm
	oStyle.Width = 21000
	oStyle.Height = 29700
	oStyle.isLandscape = False
End Sub

