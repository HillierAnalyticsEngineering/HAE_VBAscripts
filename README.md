# HAE_VBAscripts
A collection of useful, modular scripts, class modules, and modules that may be implemented in VBA projects to abstract away complexity or automate processes.

### 6/26/2022:
##cSaveState.vb 
    a script intended for use as a VBA class module. This class abstracts away the need to work with filesystem objects
    uses pipe-delimited lines of plain text to perform CRUD operations. Originally was designed to persist GUI object property settings for user forms,
    but likely has many other uses.
    
##TestFileIO 
    a script used to demonstrate and test the functionality of the cSaveState class module.
