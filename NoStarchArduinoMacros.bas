Attribute VB_Name = "NoStarchArduinoMacros"
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

  ' colors according to Arduino IDE. These are not changed in the NS Styles
  colors(1) = RGB(211, 84, 0)     'orange
  colors(2) = RGB(211, 84, 0)     'orange
  colors(3) = RGB(114, 142, 0)    'green
  colors(4) = RGB(0, 151, 156)    'blue
  colors(5) = RGB(102, 102, 102)  'grey
  
End Sub

Public Sub NSarduinoCodePretty()
Attribute NSarduinoCodePretty.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.arduinoCodePretty"
'
' No Starch arduinoCodePretty Macro
' Highlights all text and baselines the style as CodeB
' Searches for keywords based on Arduino IDE and changes the styles accordingly
'
' Search for multi-line comments does not work fully. May need to still go back and
' manually clean up multi-line comments to Arduino Grey
'

    If (Selection.Range = "") Then
        MsgBox ("Please select the code you wish to format.")
        Exit Sub
    End If

    setup
    
    Application.ScreenUpdating = False
       
    Dim formatRange As Range
    Set formatRange = Selection.Range
    
    Selection.Style = "CodeB"
    
    For i = 0 To UBound(keyword1)
        changeStyle r:=formatRange, searchStr:=keyword1(i), StyleName:="Arduino Orange"
    Next i
    
    For i = 0 To UBound(keyword2)
        changeStyle r:=formatRange, searchStr:=keyword2(i), StyleName:="Arduino Orange"
    Next i
        
    For i = 0 To UBound(keyword3)
        changeStyle r:=formatRange, searchStr:=keyword3(i), StyleName:="Arduino Olive Green"
    Next i
        
    For i = 0 To UBound(literal1)
        changeStyle r:=formatRange, searchStr:=literal1(i), StyleName:="Arduino Dark Teal"
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
            Selection.Style = ActiveDocument.Styles("Arduino Grey")
            Selection.Move Unit:=wdLine, Count:=1
        End If
    Loop Until (Selection.Find.Found = False)
    
End Sub

Private Sub findMultiLineComments(r As Range)

    r.Select
    Do
        Selection.Find.Text = "/*"
        Selection.Find.Forward = True
        Selection.Find.Execute
        If (Selection.Find.Found) Then
            Selection.MoveEndUntil Cset:="/"
            Selection.MoveEnd Unit:=wdCharacter, Count:=1
            Selection.Style = ActiveDocument.Styles("Arduino Grey")
            Selection.Move Unit:=wdLine, Count:=1
        End If
    Loop Until (Selection.Find.Found = False)
    
End Sub
Private Sub changeStyle(r As Range, ByVal searchStr As String, StyleName As String)
    
    r.Select
    Do
        Selection.Find.Text = searchStr
        Selection.Find.Forward = True
        Selection.Find.MatchCase = True
        Selection.Find.Execute
        If (Selection.Find.Found) Then
            Selection.Style = ActiveDocument.Styles(StyleName)
            Selection.Move Unit:=wdCharacter, Count:=1
        End If
    Loop Until (Selection.Find.Found = False)
End Sub
