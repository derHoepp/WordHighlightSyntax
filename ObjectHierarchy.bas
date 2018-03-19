Attribute VB_Name = "ObjectHierarchy"
Option Explicit
'
'Classes and Objects and public Member
'asdf
'Highlighter
 '| (BaseClass)
 '|-#CreateCharacterStyles
 '|-CodeSections
 '|  |#Add (CodeSection)
 '|  |-Count
 '|  |-Item (CodeSection)
 '|  \-CodeSection
 '|     |-HasChanged (Boolean)
 '|     |-Paragraphs
 '|     |-#ParseParagraphs
 '|     \-Lines
 '|        |-Line
 '|        |  |#ParseParagraph (Word.Paragraph)
 '|        |  \-Keywords
 '|        |     |-#Add (Keyword)
 '|        |     |-Count(Long)
 '|        |     |-Item (Keyword)
 '|        |     \-Keyword
 '|        |        |-Range (Word.Range)
 '|        |        |-Tag (String)
 '|        |        \-KeywordType (String)
 '|        |#Add (Line)
 '|        |-Item (Line)
 '|        \-Count (Line)
 '|-KeyWordDictionary
 '|  |-#InitFromFile (String)
 '|  |-GetCategoryFromKeyword (String)
 '|  \-GetCategories (Variant())
 '|-Document (Word.Document)
 '|-#ParseDocument
