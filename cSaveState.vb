Option Explicit

' Class saveState

Private fPath As String
Private readData As Collection
Private readKey As Collection
Private str As String
Private x As Variant
Private i As Integer
Private timeStamp As String 'for Log, always Format(Now, "yyyy-mm-dd HH:mm:ss")
Private setting As String
Private objFSO
Private objTF
Private objTS

Private Sub Class_Initialize()

    Set objFSO = CreateObject("scripting.filesystemobject")
    Set readData = New Collection
    Set readKey = New Collection

End Sub

Public Property Get FilePath() As String
    FilePath = fPath
End Property

Public Property Let FilePath(ByVal value As String)
    fPath = value
End Property

Public Function CreateFile(Optional ByVal setting_ As String)
    
    'default setting generates it's own relative fpath in xlsm dir so user need not specify
    '
    If setting_ = "Default" Or setting_ = "default" Then
        fPath = ".\\saveStateFile.SAVESTATE"
        setting = "default"
    End If
    
    
    'log setting creates a logfile and gives access to .CreateLog()
    'generates it's own relative fpath in xlsm dir so user need not specify
    '   -this is the only way for class to get "log" setting (by design)
    '
    If setting_ = "Log" Or setting_ = "log" Then
        fPath = ".\\application.LOG"
        setting = "log"
    End If
    
    
    'create new plain text file
    '
    Set objTF = objFSO.CreateTextFile(fPath, True, False)
    objTF.Close

End Function

Public Function Record(ByVal key As String, ParamArray vals() As Variant)

    'build a string with key and paramarray values
    'Init String
    'Capture first value
    'Append remaining values, pipe delimited (start i to 1 instead of 0)
    'prepend key to values, pipe delimited
    '
    str = ""
    str = str + vals(0)
    For i = 1 To UBound(vals, 1)
        str = str + "|" + vals(i)
    Next i
    str = key + "|" + str
    
    
    'open file with reading ability only
    'read all keys for matching by key
    'close the instance of read mode (1)
    '
    Set objTS = objFSO.OpenTextFile(fPath, 1)
    Do While objTS.AtEndofStream <> True
        readKey.Add (Split(objTS.readLine, "|")(0))
    Loop
    objTS.Close
    
    
    'Open text file in append mode (8)
    'check if file is empty - if so safe to write the line, close the file, and exit function
    'check existing keys - if key exists, close file and exit function - else append line
    '   -to update use saveFile.Update
    '
    Set objTS = objFSO.OpenTextFile(fPath, 8)
    If readKey.Count = 0 Then
        objTS.WriteLine str
        objTS.Close
        Exit Function
    End If
    For Each x In readKey
        If StrComp(x, key, vbTextCompare) = 0 Then
            objTS.Close
            Exit Function
        End If
    Next x
    
    
    'Passed, write the line & close
    '
    objTS.WriteLine str
    objTS.Close
    
    
    'reset collection
    Set readKey = New Collection

End Function

Public Function Update(ByVal key As String, ParamArray vals() As Variant)

    'note, this data structure is not meant for hundreds of sheets, only a handful!
    
    'Make sure collections are empty (literally does not work without this at top & bottom)
    '
    Set readData = New Collection
    Set readKey = New Collection
    

    'open file with reading ability only
    'read all lines at O(N) time complexity using do-while loop
    'stores lines in a collection of lines
    'read all keys for matching by key in collection of keys
    'close the instance of read mode (1)
    '
    Set objTS = objFSO.OpenTextFile(fPath, 1)
    i = 1
    Do While objTS.AtEndofStream <> True
        readData.Add (objTS.readLine)
        readKey.Add (Split(readData(i), "|")(0))
        i = i + 1
    Loop
    objTS.Close
    
    
    'build a string with key and paramarray values
    'Init String
    'Capture first value
    'Append remaining values, pipe delimited (start i to 1 instead of 0)
    'prepend key to values, pipe delimited
    '
    str = ""
    str = str + vals(0)
    For i = 1 To UBound(vals, 1)
        str = str + "|" + vals(i)
    Next i
    str = key + "|" + str
    
    
    'find data to update by key and replace with key and new values
    '
    i = 1
    For Each x In readKey
        If StrComp(x, key, vbTextCompare) = 0 Then
            readData.Remove i
            readData.Add str
            Exit For
        End If
        i = i + 1
    Next x
    
    
    'open file with overwrite mode (2)
    'write updated lines back to file
    'close instance of overwrite mode (2)
    '
    Set objTS = objFSO.OpenTextFile(fPath, 2)
    For Each x In readData
        objTS.WriteLine x
    Next x
    objTS.Close
    
    
    'reset collection
    Set readData = New Collection
    Set readKey = New Collection
    
End Function

Public Function Read(ByVal key As String) As String
    
    'open file with reading ability only
    'read all lines at O(N) time complexity using do-while loop
    'stores lines in a collection of lines
    'read all keys for matching by key in collection of keys
    'close the instance of read mode (1)
    '
    Set objTS = objFSO.OpenTextFile(fPath, 1)
    i = 1
    Do While objTS.AtEndofStream <> True
        readData.Add (objTS.readLine)
        readKey.Add (Split(readData(i), "|")(0))
        i = i + 1
    Loop
    objTS.Close
    
    
    'find data to read by key and return item with matching key
    '
    i = 1
    For Each x In readKey
        If StrComp(x, key, vbTextCompare) = 0 Then
            Read = readData(i)
            Exit Function
        End If
        i = i + 1
    Next x
    
    
    'If no Key Match
    '
    Read = "NULL"
    
    
    'reset collection
    Set readData = New Collection
    Set readKey = New Collection

End Function

Public Function Delete(ByVal key As String)

    'note, this data structure is not meant for hundreds of sheets, only a handful!
    
    'Make sure collections are empty (literally does not work without this at top & bottom)
    '
    Set readData = New Collection
    Set readKey = New Collection


    'open file with reading ability only
    'read all lines at O(N) time complexity using do-while loop
    'stores lines in a collection of lines
    'read all keys for matching by key in collection of keys
    'close the instance of read mode (1)
    '
    Set objTS = objFSO.OpenTextFile(fPath, 1)
    i = 1
    Do While objTS.AtEndofStream <> True
        readData.Add (objTS.readLine)
        readKey.Add (Split(readData(i), "|")(0))
        i = i + 1
    Loop
    objTS.Close
    
    
    'find data to update by key and replace with key and new values
    '
    i = 1
    For Each x In readKey
        If StrComp(x, key, vbTextCompare) = 0 Then
            readData.Remove i
            Exit For
        End If
        i = i + 1
    Next x
    
    
    'open file with overwrite mode (2)
    'write updated lines back to file
    'close instance of overwrite mode (2)
    '
    Set objTS = objFSO.OpenTextFile(fPath, 2)
    For Each x In readData
        objTS.WriteLine x
    Next x
    objTS.Close
    
    
    'reset collection
    Set readData = New Collection
    Set readKey = New Collection

End Function

Public Function CreateLog(ByVal modName_ As String, ParamArray vals() As Variant)

    'reserves functionality to files created with the 'log' setting via .CreateFile()
    '
    If setting <> "log" Then
        MsgBox "Logs must be created in separate file - construct new savestate object with 'log' setting"
        Exit Function
    End If
    
    
    'build a string with modName, timestamp, and paramarray values
    'Init String
    'Capture first value
    'Append remaining values, pipe delimited (start i to 1 instead of 0)
    'prepend key to values, pipe delimited
    '
    str = ""
    str = str + vals(0)
    For i = 1 To UBound(vals, 1)
        str = str + "|" + vals(i)
    Next i
    timeStamp = Format(Now, "yyyy-mm-dd HH:mm:ss")
    str = modName_ + "|" + timeStamp + "|" + str
    
    
    'Append logs to .LOG file
    Set objTS = objFSO.OpenTextFile(fPath, 8)
    objTS.WriteLine str
    objTS.Close

End Function
