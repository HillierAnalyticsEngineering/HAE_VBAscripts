Option Explicit

Sub TestFileIO()

    Dim sf As New cSaveState
    Dim log As New cSaveState
    Dim v1, v2, v3 As String
    Dim k, k2, k3 As String
    Dim str As String
    
    k = "asdf-789"
    k2 = "fdsa-987"
    k3 = "deft-567"
    v1 = "Enabled"
    v2 = CStr(2)
    v3 = CStr(188)
    
    'set filepath & create file
    With sf
        .FilePath = ".\\mySaveFile.SAVESTATE"
        .CreateFile
    End With
    
    'create log (creates it's own filepath)
    log.CreateFile "log"
    
    'add data to save file by key and log
    With sf
        .Record k, v2, v2, v1, v2
        .Record k2, v1, v2, v3, v1, v1, v1, v1, v3
        .Record k3, v3, v3, v1
    End With
    log.CreateLog "RecordVals", k, v2, v2, v1, v2
    log.CreateLog "RecordVals", k2, v1, v2, v3, v1, v1, v1, v1, v3
    log.CreateLog "RecordVals", k3, v3, v3, v1
    
    'notice these lines don't record because of violating duplicate key rule
    With sf
        .Record k, v1
        .Record k2, v1
        .Record k3, v1
    End With
    log.CreateLog "RecordVals", k, v1
    log.CreateLog "RecordVals", k2, v1
    log.CreateLog "RecordVals", k3, v1

    'read data by key
    MsgBox sf.Read(k)
    MsgBox sf.Read(k2)
    MsgBox sf.Read(k3)
    
    'update data by key, notice these overwrite prior entries
    With sf
        .Update k, v1
        .Update k2, v3
    End With
    log.CreateLog "UpdateVals", k, v1
    log.CreateLog "UpdateVals", k2, v3

    'read updated data by key
    MsgBox sf.Read(k)
    MsgBox sf.Read(k2)

    'delete data by key
    sf.Delete k2
    log.CreateLog "DeleteVals", k2

    'try to read k2 now, Not Found
    MsgBox sf.Read(k2)

    'check folder containing xlsm, should include
    'the .SAVESTATE plain text file and .LOG plain text file
    
'    Final File contain:
'    deft-567|188|188|Enabled
'    asdf-789|Enabled
'
'    Final Log should contain:
'    RecordVals|2022-06-26 10:33:02|asdf-789|2|2|Enabled|2
'    RecordVals|2022-06-26 10:33:02|fdsa-987|Enabled|2|188|Enabled|Enabled|Enabled|Enabled|188
'    RecordVals|2022-06-26 10:33:02|deft-567|188|188|Enabled
'    RecordVals|2022-06-26 10:33:02|asdf-789|Enabled
'    RecordVals|2022-06-26 10:33:02|fdsa-987|Enabled
'    RecordVals|2022-06-26 10:33:02|deft-567|Enabled
'    UpdateVals|2022-06-26 10:33:05|asdf-789|Enabled
'    UpdateVals|2022-06-26 10:33:05|fdsa-987|188
'    DeleteVals|2022-06-26 10:33:06|fdsa-987
    
End Sub
