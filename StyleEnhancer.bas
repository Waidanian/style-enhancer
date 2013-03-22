
' Visual guide for sizing the IDE window,
' so that it does not exceed 80 columns:
'     ->'     ->'     ->'     ->'     ->'     ->'     ->'     ->'     ->'   80->

' Everything gets thrown into this module for now. Appropriate separation
' will comme later.

' 2-space indents are used because some statements can get quite long.

Option Explicit

sub applyCharacterStyletoNextWord()

  ' Variable declarations
  dim theCurrentDocument As Object, currentDocText As Object
  dim styleCursor As Object

  ' Variable assignments
  theCurrentDocument = ThisComponent
  currentDocText = theCurrentDocument.Text

  ' Task
  styleCursor = currentDocText.createTextCursor
  styleCursor.gotoNextParagraph(false)
  styleCursor.gotoNextWord(false)
  styleCursor.gotoEndOfWord(true)
  styleCursor.CharStyleName = "Emphasis"

end sub

sub printCharacterStylesofParagraph()

  ' Variable declarations
  dim theCurrentDocument As Object, currentDocText As Object
  dim styleCursor As Object

  ' Variable assignments
  theCurrentDocument = ThisComponent
  currentDocText = theCurrentDocument.Text

  ' Task
  styleCursor = currentDocText.createTextCursor
  styleCursor.gotoNextParagraph(false)
  styleCursor.gotoEndOfParagraph(true)
  print styleCursor.CharStyleName

end sub

sub findCharStyleinDocument()

' Declarations:
  dim theCurrentDocument As Object
  dim mySearchOperation As Object, foundStuff As Variant
  dim appliedCharStyle(0) As New com.sun.star.beans.PropertyValue

' Assignments:
  ' Document object
  theCurrentDocument = ThisComponent

  ' This attribute allows us to check whether content has an applied style
  appliedCharStyle(0).Name = "CharStyleNames"

  ' Search operation
  mySearchOperation = theCurrentDocument.createSearchDescriptor
  mySearchOperation.SearchAttributes = appliedCharStyle()
  mySearchOperation.ValueSearch = False

  foundStuff = theCurrentDocument.findAll(mySearchOperation)

' Print Job:
print "Occurrences: " & foundStuff.Count

End Sub

sub ReplaceCharStyleInEnumParPortions()

dim TheCurrentDoc As Object
dim TextElementEnum As Object
dim TextPortionEnum As Object
dim TextElement As Object
dim TextPortion As Object

TheCurrentDoc = ThisComponent
TextElementEnum = TheCurrentDoc.Text.createEnumeration

' Loop over all text elements and operate only on paragraphs.
while TextElementEnum.hasMoreElements
  TextElement = TextElementEnum.nextElement

  if TextElement.supportsService("com.sun.star.text.Paragraph") then
    TextPortionEnum = TextElement.createEnumeration

    ' Loop over all paragraph portions and
    ' operate on the ones that have a certain style.
    ' Question: Should we not check whether
    ' the portion element “supports”
    ' the character properties? Short answer: no, doesn’t make sense.
    while TextPortionEnum.hasMoreElements
      TextPortion = TextPortionEnum.nextElement

      if TextPortion.CharStyleName = "Emphasis" then
        TextPortion.CharStyleName = "CIERL-Accentuation-legere"
      end if
    wend
  end if
wend

end sub

sub ReplaceCharStyleInStyledEnumParPortions()

dim TheCurrentDoc As Object
dim TextElementEnum As Object
dim TextPortionEnum As Object
dim TextElement As Object
dim TextPortion As Object

TheCurrentDoc = ThisComponent
TextElementEnum = TheCurrentDoc.Text.createEnumeration

while TextElementEnum.hasMoreElements
  TextElement = TextElementEnum.nextElement

  if TextElement.supportsService("com.sun.star.text.Paragraph") then

    if TextElement.ParaStyleName = "CIERL-Exergue" then
      TextPortionEnum = TextElement.createEnumeration

      while TextPortionEnum.hasMoreElements
        TextPortion = TextPortionEnum.nextElement

        if TextPortion.CharStyleName = "CIERL-Accentuation-legere" then
          TextPortion.CharStyleName = "Bold"
        end if
      wend
    end if
  end if
wend

end sub

sub ReplaceCharStyleInStyledWithStringEnumParPortions()

dim TheCurrentDoc As Object
dim TextElementEnum As Object
dim TextPortionEnum As Object
dim TextElement As Object
dim TextPortion As Object

TheCurrentDoc = ThisComponent
TextElementEnum = TheCurrentDoc.Text.createEnumeration

while TextElementEnum.hasMoreElements
  TextElement = TextElementEnum.nextElement

  if TextElement.supportsService("com.sun.star.text.Paragraph") then

    if TextElement.ParaStyleName = "CIERL-Exergue" then ' and _
    ' TextElement contains string "" then
      TextPortionEnum = TextElement.createEnumeration

      while TextPortionEnum.hasMoreElements
        TextPortion = TextPortionEnum.nextElement

        if TextPortion.CharStyleName = "CIERL-Accentuation-legere" then
          TextPortion.CharStyleName = "Bold"
        end if
      wend
    end if
  end if
wend

end sub

sub ReplaceParaStyleOnlyIfNextParaStyleIsSuch()

dim TheCurrentDoc As Object
dim TextElementEnum As Object
dim TextPortionEnum As Object
dim TextElement As Object
dim TextPortion As Object

TheCurrentDoc = ThisComponent
TextElementEnum = TheCurrentDoc.Text.createEnumeration

while TextElementEnum.hasMoreElements
  TextElement = TextElementEnum.nextElement

  if TextElement.supportsService("com.sun.star.text.Paragraph") then

    if TextElement.ParaStyleName = "CIERL-Exergue" and _
    TextElementEnum.nextElement.ParaStyleName = "CIERL-Exergue" then
      print "There are two adjacent paragraphs with the same style."
'      TextElement.ParaStyleName = "CIERL-Corps-de-texte"

    end if
  end if
wend

end sub

sub ReplaceCharStyleofStringInStyledEnumParPortions()

dim TheCurrentDoc As Object
dim TextElementEnum As Object
dim TextPortionEnum As Object
dim TextElement As Object
dim TextPortion As Object

TheCurrentDoc = ThisComponent
TextElementEnum = TheCurrentDoc.Text.createEnumeration

while TextElementEnum.hasMoreElements
  TextElement = TextElementEnum.nextElement

  if TextElement.supportsService("com.sun.star.text.Paragraph") then

    if TextElement.ParaStyleName = "CIERL-Exergue" then
      TextPortionEnum = TextElement.createEnumeration

      while TextPortionEnum.hasMoreElements
        TextPortion = TextPortionEnum.nextElement

        if TextPortion.String = "This has exergue" then
          TextPortion.CharStyleName = "Bold"
        end if
      wend
    end if
  end if
wend

end sub
