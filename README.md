# HAE_VBAscripts
A collection of useful, modular scripts, class modules, and modules that may be implemented in VBA projects to abstract away complexity or automate processes.

### 6/26/2022:

## cSaveState.bas

    A script intended for use as a VBA class module. This class abstracts away the need to work with filesystem objects
    uses pipe-delimited lines of plain text to perform CRUD operations. Originally was designed to persist GUI object 
    property settings for user forms, but likely has many other uses.
    
    Contains functionality to be created solely for logging purposes to create a plain text application.LOG file which
    can be used to track data about how proceudres were performed, or to capture object state changes for debugging or
    regulatory tracking.
    
    
    available properties:

    .FilePath  
    
            >>  Set your own file path and extension
    
    
    available methods:
    
    .CreateFile(ByVal Optional setting_ As String)  
    
            >>  Create a text file. setting_ may be "default" or "log"
                >  If setting_ not applied, .FilePath must be set on object prior to call.
            
    .Record(ByVal key As String, ParamArray vals() As Variant)  
    
            >>  Appends a pipe-delimited record to the file
                >  Key is a unique string
                >  vals() are one or more strings - pass numeric values in using CStr(numeric)
                >  if Key matches an existing key, no write occurs - use update to overwrite originally written values
            
    .Update(ByVal key As String, ParamArray vals() As Variant)  
    
            >>  Changes a pipe-delimited record by matching to Key
                >  Key must be a unique string that matches an existing key in the file
                >  vals() are one or more strings - pass numeric values in using CStr(numeric)
            
    .Read(ByVal key As String) As String  
    
            >>  Reads pipe-delimited record by matching to key and returns a string
                >>  Key must be a unique string that matches an existing key in the file
                >>  If no match is found, returns "NULL"
            
    .Delete(ByVal key As String)  
    
            >>  Removes pipe-delimited record by matching to key
                >  Key must be a unique string that matches an existing key in the file
                >  Permanently deletes record - can not be retrieved unless a backup file was created
            
    .CreateLog(ByVal modName_ As String, ParamArray vals() As Variant)  
    
            >>  Creates a pipe-delimited log record
                >  Can only be used if "log" passed to .CreateFile() to update setting_ data member
                >  No Key specified or needed
                >  Creates a timestamp in the format of Format(Now, "yyyy-mm-dd HH:mm:ss")
                >  modName_ could be a module name, a subprocedure name, or otherwise - use to identify tracked process
                >  vals() are one or more strings - pass numeric values in using CStr(numeric)
                
    .FileExists()
      
            >> Checks if a file exists at the object's filepath, returns a Boolean (True Exists, False Does not Exist)
                >  Use to avoid overwriting files constantly by making CreateFile optional based on file existance
    
## TestFileIO.bas

    A script used to demonstrate and test the functionality of the cSaveState class module.
