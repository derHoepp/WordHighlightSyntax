# Information about Class and Objecthierarchy

## Objecthierarchy
Public member and methods of the class-system. Subs are marked with a #.
```
Highlighter
 | (BaseClass)
 |-#CreateCharacterStyles
 |-#CreateCodeParagraphStyle
 |-CodeSections
 |  |#Add (CodeSection)
 |  |-Count
 |  |-Item (CodeSection)
 |  \-CodeSection
 |     |-HasChanged (Boolean)
 |     |-Paragraphs
 |     |-#ParseParagraphs
 |     \-Lines
 |        |-Line
 |        |  |#ParseParagraph (Word.Paragraph)
 |        |  \-Keywords
 |        |     |-#Add (Keyword)
 |        |     |-Count(Long)
 |        |     |-Item (Keyword)
 |        |     \-Keyword
 |        |        |-Range (Word.Range)
 |        |        |-Tag (String)
 |        |        \-KeywordType (String)
 |        |#Add (Line)
 |        |-Item (Line)
 |        \-Count (Line)
 |-KeyWordDictionary
 |  |-#InitFromFile (String)
 |  |-GetCategoryFromKeyword (String)
 |  \-GetCategories (Variant())
 |-Document (Word.Document)
 \-#ParseDocument
```
