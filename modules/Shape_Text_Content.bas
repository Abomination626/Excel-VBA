Attribute VB_Name = "ShapeText"
Sub Shape_text_manipulation()
'******************************************
'THIS SNIPPET SHOWS HOW TO MANAGE TEXT CONTENTS
'CAN BE USEFUL FOR DOCUMENT TRANSLATION THAT HAVE
'SHAPES WITH TEXT
'******************************************
Dim shp As Shape
i = 0
For Each shp In ActiveSheet.Shapes
  Debug.Print shp.Name
  shp.TextFrame.Characters.text = "text " & i
i = i + 1
Next shp

End Sub
