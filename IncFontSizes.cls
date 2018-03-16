VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''
' *** IncFontSizes v1.0.0 ***
' by Craig Lavender
'
' IncFontSizes proportionately increases the font size of selected
' text by a given percentage.
'
' to use select text and run macro


Public Sub IncFontSizes()
  Dim rngChar As Range
  Dim IncInput As String
  Dim CharCount As Long
  Dim FontSize As Variant
  Dim NewFontSize As Variant
  Dim FontInc As Variant
  Dim Link As Hyperlink
  Dim rngLink As Range
  Dim HyperFont As Integer
  
  CharCount = 0
  ' *** I was using below section before I figured out how to set size of hyperlink font. ***
'  If Selection.Hyperlinks.Count > 0 Then
'      Result = MsgBox("Selection contains Hyperlinks!" & _
'               Chr(13) & "These will not be" & _
'               Chr(13) & "automatically resized!", 48, "HYPERLINK WARNING")
'  End If
  IncInput = InputBox(Prompt:="Input font percentage increase required " & _
              wdChrLf & "for selected text. " & _
              wdChrLf & "Selection must be contiguous.", _
 Title:="Font % Increase", Default:="Enter percentage value")

 If Not IsNumeric(FontInc) Then
   MsgBox "You can only enter a number in this field"
 Else
   FontInc = Val(IncInput)
   If FontInc < -50 Or FontInc > 500 Then
     MsgBox "Percentage must be between -50 and 500!"
   Else
     ' MsgBox "FontInc= " & FontInc
     ' Do stuff
     ' Application.Cursor = xlWait
     System.Cursor = wdCursorWait
     ' Turning off screen updates greatly improves VBA macro speeds
     Application.ScreenUpdating = False
     ' For Each rngChar In ActiveDocument.Characters
     For Each rngChar In Selection.Characters
       CharCount = CharCount + 1
       If rngChar.Hyperlinks.Count < 1 Then
         FontSize = rngChar.Font.Size
         NewFontSize = FontSize * (1 + (FontInc / 100))
         rngChar.Font.Size = NewFontSize
         ' I can't find the syntax to read Hyperlink sizes,
         ' so I to keep last char size to use if we come across a Hyperlink
         HyperFont = CInt(NewFontSize)
       Else
         rngChar.Hyperlinks(1).Range.Font.Size = HyperFont  ' Set Hyperlink font size to same as last text char.
       End If
     Next rngChar
     ' Turn screen updates back on
     Application.ScreenUpdating = True
     System.Cursor = wdCursorNormal
   End If
 End If
  
  
  
End Sub


