Attribute VB_Name = "ArduinoCodePretty"
Dim colors(1 To 5) As Long
Dim keyword1 As Variant
Dim keyword2 As Variant
Dim keyword3 As Variant
Dim literal1 As Variant

Private Sub setup()
  keyword1 = Array("Serial", "Serial1", "Serial2", "Serial3", "SerialUSB", "Keyboard", "Mouse")
  
  keyword2 = Array("abs", "acos", "asin", "atan", "atan2", "ceil", "constrain", "cos", "degrees", _
    "exp", "floor", "log", "map", "max", "min", "radians", "random", "randomSeed", "round", "sin", _
    "sq", "sqrt", "tan", "pow", "bitRead", "bitWrite", "bitSet", "bitClear", "bit", "highByte", _
    "lowByte", "analogReference", "analogRead", "analogReadResolution", "analogWrite", _
    "analogWriteResolution", "attachInterrupt", "detachInterrupt", "digitalPinToInterrupt", _
    "delay", "delayMicroseconds", "digitalWrite", "digitalRead", "interrupts", "millis", "micros", _
    "noInterrupts", "noTone", "pinMode", "pulseIn", "pulseInLong", "shiftIn", "shiftOut", "tone", _
    "yield", "Stream", "begin", "end", "peek", "read", "print", "println", "available", "availableForWrite", _
    "flush", "setTimeout", "find", "findUntil", "parseInt", "parseFloat", "readBytes", "readBytesUntil", _
    "readString", "readStringUntil", "trim", "toUpperCase", "toLowerCase", "charAt", "compareTo", "concat", _
    "endsWith", "startsWith", "equals", "equalsIgnoreCase", "getBytes", "indexOf", "lastIndexOf", "length", _
    "replace", "setCharAt", "substring", "toCharArray", "toInt", "press", "release", "releaseAll", "accept", _
    "click", "move", "isPressed", "isAlphaNumeric", "isAlpha", "isAscii", "isWhitespace", "isControl", "isDigit", _
    "isGraph", "isLowerCase", "isPrintable", "isPunct", "isSpace", "isUpperCase", "isHexadecimalDigit")

  keyword3 = Array("break", "case", "override", "final", "continue", "default", "do", "else", "for", "if", "return", _
    "goto", "switch", "throw", "try", "while", "setup", "loop", "export", "not", "or", "and", "xor", "#include", _
    "#define", "#elif", "#else", "#error", "#if", "#ifdef", "#ifndef", "#pragma", "#warning")


  literal1 = Array("HIGH", "LOW", "INPUT", "INPUT_PULLUP", "OUTPUT", "DEC", "BIN", "HEX", "OCT", "PI", "HALF_PI", "TWO_PI", _
    "LSBFIRST", "MSBFIRST", "CHANGE", "FALLING", "RISING", "DEFAULT", "EXTERNAL", "INTERNAL", "INTERNAL1V1", "INTERNAL2V56", _
    "LED_BUILTIN", "LED_BUILTIN_RX", "LED_BUILTIN_TX", "DIGITAL_MESSAGE", "FIRMATA_STRING", "ANALOG_MESSAGE", "REPORT_DIGITAL", _
    "REPORT_ANALOG", "SET_PIN_MODE", "SYSTEM_RESET", "SYSEX_START", "auto", "int8_t", "int16_t", "int32_t", "int64_t", "uint8_t", _
    "uint16_t", "uint32_t", "uint64_t", "char16_t", "char32_t", "operator", "enum", "delete", "bool", "boolean", "byte", "char", _
    "const", "false", "float", "double", "null", "NULL", "int", "long", "new", "private", "protected", "public", "short", "signed", _
    "static", "volatile", "String", "void", "true", "unsigned", "word", "array", "sizeof", "dynamic_cast", "typedef", "const_cast", _
    "struct", "static_cast", "union", "friend", "extern", "class", "reinterpret_cast", "register", "explicit", "inline", "_Bool", _
    "complex", "_Complex", "_Imaginary", "atomic_bool", "atomic_char", "atomic_schar", "atomic_uchar", "atomic_short", "atomic_ushort", _
    "atomic_int", "atomic_uint", "atomic_long", "atomic_ulong", "atomic_llong", "atomic_ullong", "virtual", "PROGMEM")


  colors(1) = RGB(211, 84, 0)     'orange
  colors(2) = RGB(211, 84, 0)     'orange
  colors(3) = RGB(114, 142, 0)    'green
  colors(4) = RGB(0, 151, 156)    'blue
  colors(5) = RGB(102, 102, 102)  'grey
  
End Sub

Public Sub ArduinoCodePretty()
Attribute ArduinoCodePretty.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.arduinoCodePretty"
'
' arduinoCodePretty Macro
'
'
    If (Selection.Range = "") Then
        MsgBox ("Please select the code you wish to format.")
        Exit Sub
    End If
    
    setup
    
    Application.ScreenUpdating = False
    Selection.Font.Name = "Courier New"
    Selection.Font.Size = 9
    Dim formatRange As Range
    
    Set formatRange = Selection.Range
    For i = 0 To UBound(keyword1)
        changeColor r:=formatRange, searchStr:=keyword1(i), color:=colors(1)
    Next i
    
    For i = 0 To UBound(keyword2)
        changeColor r:=formatRange, searchStr:=keyword2(i), color:=colors(2)
    Next i
        
    For i = 0 To UBound(keyword3)
        changeColor r:=formatRange, searchStr:=keyword3(i), color:=colors(3)
    Next i
        
    For i = 0 To UBound(literal1)
        changeColor r:=formatRange, searchStr:=literal1(i), color:=colors(4)
    Next i
       
    findSingleComments r:=formatRange
    findMultiLineComments r:=formatRange
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub findSingleComments(r As Range)

    r.Select
    
    Do
        Selection.Find.Text = "//"
        Selection.Find.Forward = True
        Selection.Find.Execute
        If (Selection.Find.Found) Then
            Selection.MoveEndUntil Cset:=vbCrLf
            Selection.Font.color = colors(5)  'set to grey font color
            Selection.Move Unit:=wdLine, count:=1
        End If
    Loop Until (Selection.Find.Found = False)
    
End Sub

Private Sub findMultiLineComments(r As Range)
    Dim endtoken As Range
    
    r.Select
    Set endtoken = Selection.Range
        r.Find.Text = "/*"
        r.Find.Forward = True
        endtoken.Find.Text = "*/"
        endtoken.Find.Forward = True
    
    Do
        r.Find.Execute
        If (r.Find.Found) Then
            endtoken.Find.Execute
            If (endtoken.Find.Found) Then
                ActiveDocument.Range(r.Start, endtoken.End).Select
                Selection.Font.color = colors(5)
                endtoken.Move Unit:=wdCharacter, count:=1
                r.Move Unit:=wdCharacter, count:=1
            End If
        End If
    Loop Until (r.Find.Found = False)
    
End Sub
Private Sub changeColor(r As Range, ByVal searchStr As String, color As Long)
Attribute changeColor.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    r.Select
    Do
        With Selection.Find
            .Text = searchStr
            .Forward = True
            .MatchCase = True
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        
        If (Selection.Find.Found) Then
            Selection.Font.color = color
            Selection.Move Unit:=wdCharacter, count:=1
        End If
    Loop Until (Selection.Find.Found = False)
End Sub
Public Sub formatLineNums()
Attribute formatLineNums.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    
    Dim linNums As String
    Dim formatRange As Range
    Dim numLines As Integer
    
    If (Selection.Range = "") Then
        MsgBox ("Please select the text for the code you want to format.")
        Exit Sub
    End If
    
    Set formatRange = Selection.Range

    formatRange.Select
    
    'Find the total number of lines of code
    numLines = formatRange.ComputeStatistics(wdStatisticLines)
    
    'Cut the Selection to paste into a table
    Selection.Cut
    
    'Create a table with two columns and format the widths to accomodate line numbers
    'on left side and up to 80 characters of fixed width font on the right
    ActiveDocument.Tables.Add Range:=Selection.Range, numRows:=1, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
        .Columns(1).SetWidth ColumnWidth:=30, RulerStyle:=wdAdjustFirstColumn
        .Columns(2).SetWidth ColumnWidth:=456, RulerStyle:=wdAdjustFirstColumn
    End With
    
    Selection.Tables(1).Cell(1, 2).Select
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.Tables(1).Cell(1, 1).Select
    
    'Create a string containing all of the line numbers
    For i = 1 To numLines
        linNums = linNums + Str(i) + ":" + vbCrLf
    Next i
    
    'Fill left-most cell with line numbers
    Selection.Text = linNums
    Selection.Font.Name = "Courier New"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    With Selection.Cells(1)
        .LeftPadding = InchesToPoints(0.01)
        .RightPadding = InchesToPoints(0.01)
    End With
    
    ActiveDocument.Tables(1).Select
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1.15)
    Selection.Move Unit:=wdCharacter, count:=1
End Sub
