# IncFontSizes
WORD VBA to proportionately increase size of font by a percentage for selected text.
Created this for my use, but also used as my first Github exercise.
Developed on Win 7, and WORD 2010.
Scenario: you've received a WORD file with numerous font sizes, but you want to increase/decrease them all proportionately.
Usage: With IncFontSizes, select all, or only a section, and run the macro. You will be prompted for a percentage change +/-.
The macro then runs through the file char by char. Hyperlinks have presented a problem for me to find the syntax to read the
current font size, so I have adopted the strategy of using whatever the immediate previous char size was. This seems to work well.
