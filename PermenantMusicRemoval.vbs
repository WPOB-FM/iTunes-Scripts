' =================
' KillSelectedFiles
' =================
' ====================================
' Made By Joseph Liszovics for WPOB-FM
' ====================================
' ===========
' Description
' ===========
' Delete selected files from iTunes and the hard drive, irrespective of location
' Recycle local files, network files will be deleted regardless

' Written partly in response to this thread on Apple Support Communities
' https://discussions.apple.com/message/18710163#18710163


' =============================
' Declare constants & variables
' =============================
' Variables for common code
Option Explicit	        ' Declare all variables before use
Dim Intro,Outro,Check   ' Manage confirmation dialogs
Dim PB,Prog,Debug       ' Control the progress bar
Dim Clock,T1,T2,Timing  ' The secret of great comedy
Dim Named,Source        ' Control use on named playlist
Dim Playlist,List       ' Name for any generated playlist, and the object itself
Dim iTunes              ' Handle to iTunes application
Dim Tracks              ' A collection of track objects
Dim Count               ' The number of tracks
Dim D,M,P,S,U,V         ' Counters
Dim nl,tab              ' New line/tab strings
Dim IDs                 ' A dictionary object used to ensure each object is processed once
Dim Rev                 ' Control processing order, usually reversed
Dim Quit                ' Used to abort script
Dim Title,Summary       ' Text for dialog boxes
Dim Tracing             ' Display/suppress tracing messages

' Values for common code

Const Kimo=False        ' True if script expects "Keep iTunes Media folder organised" to be disabled
Const Min=1             ' Minimum number of tracks this script should work with
Const Max=0             ' Maximum number of tracks this script should work with, 0 for no limit
Const Warn=500          ' Warning level, require confirmation for processing above this level
Intro=True              ' Set false to skip initial prompts, avoid if non-reversible actions
Outro=True              ' Produce summary report
Check=True              ' Track-by-track confirmation, can be set during Intro
Prog=True               ' Display progress bar, may be disabled by UAC/LUA settings
Debug=True              ' Include any debug messages in progress bar
Timing=True             ' Display running time in summary report
Named=False             ' Force script to process specific playlist rather than current selection or playlist
Source=""               ' Named playlist to process, use "Library" for entire library
Playlist=""             ' Used if this script creates a playlist
Rev=False               ' Control processing order, usually reversed
Debug=True              ' Include any debug messages in progress bar
Tracing=True            ' Display tracing message boxes


' Additional variables for this particular script
Dim UseTrash            ' Attempt to send local deleted files to trash
Dim FSO                 ' Handle to FileSystemObject
Dim Reg                 ' Handle to Registry object
Dim SH                  ' Handle to Shell application

' Initialise variables for this particular script
UseTrash=True           ' Attempt to send local deleted files to trash, false to delete permanently
Playlist=""
Set FSO=CreateObject("Scripting.FileSystemObject")
Set SH=CreateObject("Shell.Application")
Set Reg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 	' Use . for local computer, otherwise could be computer name or IP address

Title="Kill Selected Files"
Summary="Delete selected files from iTunes and the hard drive, irrespective of location."
If UseTrash Then 
  Summary=Summary & vbCrLF & vbCrLf & "Locally stored files will be sent to the recycle bin where supported."
Else
  Summary=Summary & vbCrLF & vbCrLf & "Files will be deleted without using the recycle bin."
End If


' ============
' Main program
' ============

GetTracks               ' Set things up
ProcessTracks 	        ' Main process 
Results                 ' Summary

' ===================
' End of main program
' ===================


' ===============================
' Declare subroutines & functions
' ===============================


' Note: The bulk of the code in this script is concerned with making sure that only suitable tracks are processed by
'       the following module and supporting numerous options for track selection, confirmation, progress and results.


' Recycle file and delete from iTunes

Sub Action(T)
  StartEvent            ' Time potentially slow event
    ' Comment out next two lines to test script without removing tracks
    If T.Kind=1 Then If T.Location<>"" Then Recycle Replace(T.Location,"/","\")
    DeleteTrack T
    U=U+1
  StopEvent             ' Show event time
End Sub


' Delete a track from iTunes, leave the file

Sub DeleteTrack(T)
  Dim H,L,O
  H=iTunes.ITObjectPersistentIDHigh(T)
  L=iTunes.ITObjectPersistentIDLow(T)
  Set O=iTunes.LibraryPlaylist.Tracks.ItemByPersistentID(H,L)
  O.Delete
End Sub


' Custom info message for progress bar

Function Info(T)
  Dim A,B
  A="Unknown Artist"
  B="Unknown Album"
  On Error Resume Next
  With T
    A=.AlbumArtist: If A="" Then A=.Artist : If A="" Then A="Unknown Artist"
    B=.Album : If B="" Then B="Unknown Album"
    Info="Deleting: " & A & " - " & B & " - " & .Name
  End With
End Function


' Custom prompt for track-by-track confirmation
' Modified 2011-11-13
Function Prompt(T)
  Prompt="Delete file:" & nl & nl & T.Location & " ?"
End Function


' Recycled from http://gallery.technet.microsoft.com/scriptcenter/191eb207-3a7e-4dbc-884d-5f4498440574
' Modified to recursively remove any emptied folders. Rewritten to simplify and use global objects/declarations
' Needs FSO,Reg,SH objects. If UseTrash is false delete directly without attempting to recycle.

' Send file or folder to recycle bin, return status
' Modified 2015-01-18
Function Recycle(FilePath)
  Const HKEY_CURRENT_USER=&H80000001 
  Const KeyPath="Software\Microsoft\Windows\CurrentVersion\Explorer" 
  Const KeyName="ShellState" 
  Dim File,FileName,Folder,FolderName,I,Parent,State,Value,Verb,Q,R
  Recycle=False
  If Not(FSO.FileExists(FilePath) Or FSO.FolderExists(FilePath)) Then Exit Function     ' Can't delete something that isn't there
  If  UseTrash Then
    ' Make sure recycle bin properties are set to NOT display request for delete confirmation 
    Reg.GetBinaryValue HKEY_CURRENT_USER,KeyPath,KeyName,Value			' Get current shell state 
    State=Value(4)	 							' Preserve current option
    Value(4)=39									  ' Set new option 
    Reg.SetBinaryValue HKEY_CURRENT_USER,KeyPath,KeyName,Value			' Update shell state
   
    ' Use the Shell to send the file to the recycle bin 
    FileName=FSO.GetFileName(FilePath)
    FolderName=FSO.GetParentFolderName(FilePath)
    Set Folder=SH.NameSpace(FolderName)
    Set File=Folder.ParseName(FileName)

    If Not File Is Nothing Then  
      'File.InvokeVerb("&Delete")	' Delete file, sending to recycle bin - fails for Vista/Windows 7
      I=File.Verbs.Count          ' Use DoIt instead of InvokeVerb - http://forums.wincustomize.com/322016
      Do
        I=I-1
        Verb=Replace(LCase(File.Verbs.Item(I).Name),"&","")
        ' Add lower case localised words for delete here, separated by |
        If Instr("|delete|löschen|verwijderen|","|" & Verb & "|") Then
          On Error Resume Next          ' Trap potential error
          File.Verbs.Item(I).DoIt()     ' Possible error generating line
          ' Err.Raise 1,Title,"Test"    ' Test error handler
          If Err.Number<>0 Then         ' Handle any error
            Q="An error occurred recycling a file." & nl
            Q=Q & nl & "Script:" & tab & WScript.ScriptFullName
            Q=Q & nl & "Error:" & tab & Err.Description
            Q=Q & nl & "Number:" & tab & "&" & Right("0000000" & Hex(Err.Number),8)
            Q=Q & nl & "Source:" & tab & Err.Source
            Q=Q & nl & "Path:" & tab & Wrap(FilePath,50,"\",1)
            Q=Q & nl & nl & "Press Cancel to abort this run."
            R=MsgBox(Q,vbCritical+vbOKCancel,Title)
            If R=vbCancel Then Quit=True
            Err.Clear
          End If                        'End of conditional block
          On Error Goto 0               'Reset default error handler
          Exit Do
        End If
      Loop Until I=0
      If I=0 Then             ' Assume non-English settings and predict which verb is delete
        I=File.Verbs.Count-3  ' Because delete should be the third item from the end of the list, ignoring separators
        Verb=Replace(LCase(File.Verbs.Item(I).Name),"&","")
        Trace Null,"Verb for delete not recognized" & nl & nl & "I=" & I & ", Count=" & File.Verbs.Count & ", Verb=" & Verb
        Tracing=False
        File.Verbs.Item(I).DoIt()
      End If
    End If
  Else    ' Delete via FSO instead of Shell
    FolderName=FSO.GetParentFolderName(FilePath)
    On Error Resume Next                ' iTunes operation with error handling
    FSO.DeleteFile FilePath,True        ' Possible error generating code
    ' Err.Raise 1,Title,"Test"          ' Test error handler
    If Err.Number<>0 Then               ' Handle any error
      Q="An error occurred deleting a file." & nl
      Q=Q & nl & "Script:" & tab & WScript.ScriptFullName
      Q=Q & nl & "Error:" & tab & Err.Description
      Q=Q & nl & "Number:" & tab & "&" & Right("0000000" & Hex(Err.Number),8)
      Q=Q & nl & "Source:" & tab & Err.Source
      Q=Q & nl & "Path:" & tab & Wrap(FilePath,50,"\",1)
      Q=Q & nl & nl & "Press Cancel to abort this run."
      R=MsgBox(Q,vbCritical+vbOKCancel,Title)
      If R=vbCancel Then Quit=True
      Err.Clear
    End If                              'End of conditional block
    On Error Goto 0                     'Reset default error handler
  End If
  If FSO.FileExists(FilePath) Then
    MsgBox "There was a problem deleting the file:" & nl & FilePath,vbCritical,Title
  Else
    Recycle=True
    ' Delete folder using FileSystem if now empty, repeat for parent folders
    Set Folder=FSO.GetFolder(FolderName)
    While Folder.Files.Count=0 And Folder.SubFolders.Count=0
      Set Parent=Folder.ParentFolder
      Folder.Delete
      Set Folder=Parent
    Wend
  End If

  If UseTrash Then                ' Restore the user's property settings for the Recycle Bin 
    Value(4)=State								' Restore option
    Reg.SetBinaryValue HKEY_CURRENT_USER,KeyPath,KeyName,Value			' Update shell state
  End If

End Function


' Report results...

Sub Results
  If Not Outro Then Exit Sub
  Dim T
  If Quit Then T="Script aborted!" & nl & nl Else T=""
  T=T & GroupDigits(P) & " track" & Plural(P,"s","")
  If P<Count Then T=T & " of " & count
  T=T & Plural(P," were"," was") & " processed of which " & nl
  If U>0 Then
    T=T & GroupDigits(U) & Plural(U," were"," was") & " deleted"
    If (S>0)+(M>0)<-1 Then
      T=T & "," & nl
    ElseIf (S>0)+(M>0)=-1 Then
      T=T & " and" & nl
    End If
  End If
  If S>0 Then
    T=T & GroupDigits(S) & Plural(S," were"," was") & " skipped"
    If M>0 Then T=T & " and" & nl
  End If
  If M>0 Then T=T & GroupDigits(M) & Plural(M," were"," was") & " missing"
  T=T & "."
  If Timing Then 
    T=T & nl & nl
    If Check Then T=T & "Processing" Else T=T & "Running"
    T=T & " time: " & FormatTime(Clock)
  End If
  MsgBox T,vbInformation,Title
End Sub


' Custom status message for progress bar
' Modified 2012-06-22
Function Status(N)
  Status="Processing " & GroupDigits(N) & " of " & Count
End Function


' Test for tracks which can be usefully updated

Function Updateable(T)
  Updateable=True
End Function


' ============================================
' Reusable Library Routines for iTunes Scripts
' ============================================



' Format time interval from x.xxx seconds to hh:mm:ss
Function FormatTime(T)
  If T<0 Then T=T+86400         ' Watch for timer running over midnight
  If T<2 Then
    FormatTime=FormatNumber(T,3) & " seconds"
  ElseIf T<10 Then
    FormatTime=FormatNumber(T,2) & " seconds"
  ElseIf T<60 Then
    FormatTime=Int(T) & " seconds"
  Else
    Dim H,M,S
    S=T Mod 60
    M=(T\60) Mod 60             ' \ = Div operator for integer division
    'S=Right("0" & (T Mod 60),2)
    'M=Right("0" & ((T\60) Mod 60),2)  ' \ = Div operator for integer division
    H=T\3600
    If H>0 Then
      FormatTime=H & Plural(H," hours "," hour ") & M & Plural(M," mins"," min")
      'FormatTime=H & ":" & M & ":" & S
    Else
      FormatTime=M & Plural(M," mins "," min ") & S & Plural(S," secs"," sec")
      'FormatTime=M & " :" & S
      'If Left(FormatTime,1)="0" Then FormatTime=Mid(FormatTime,2)
    End If
  End If
End Function


' Initialise track selections, quit script if track selection is out of bounds or user aborts
Sub GetTracks
  Dim Q,R
  ' Initialise global variables
  nl=vbCrLf : tab=Chr(9) : Quit=False
  M=0 : P=0 : S=0 : U=0 : V=0
  ' Initialise global objects
  Set iTunes=CreateObject("iTunes.Application")
  Set Tracks=iTunes.SelectedTracks      ' Get current selection
  If iTunes.BrowserWindow.SelectedPlaylist.Source.Kind<>1 And Source="" Then Source="Library" : Named=True      ' Ensure section is from the library source
  'If iTunes.BrowserWindow.SelectedPlaylist.Name="Ringtones" And Source="" Then Source="Library" : Named=True    ' and not ringtones (which cannot be processed as tracks???)
  If iTunes.BrowserWindow.SelectedPlaylist.Name="Radio" And Source="" Then Source="Library" : Named=True        ' or radio stations (which cannot be processed as tracks)
  If iTunes.BrowserWindow.SelectedPlaylist.Name=Playlist And Source="" Then Source="Library" : Named=True       ' or a playlist that will be regenerated by this script
  If Named Or Tracks Is Nothing Then    ' or use a named playlist
    If Source<>"" Then Named=True
    If Source="Library" Then            ' Get library playlist...
      Set Tracks=iTunes.LibraryPlaylist.Tracks
    Else                                ' or named playlist
      On Error Resume Next              ' Attempt to fall back to current selection for non-existent source
      Set Tracks=iTunes.LibrarySource.Playlists.ItemByName(Source).Tracks
      On Error Goto 0
      If Tracks is Nothing Then         ' Fall back
        Named=False
        Source=iTunes.BrowserWindow.SelectedPlaylist.Name
        Set Tracks=iTunes.SelectedTracks
        If Tracks is Nothing Then
          Set Tracks=iTunes.BrowserWindow.SelectedPlaylist.Tracks
        End If
      End If
    End If
  End If  
  If Named And Tracks.Count=0 Then      ' Quit if no tracks in named source
    If Intro Then MsgBox "The playlist " & Source & " is empty, there is nothing to do.",vbExclamation,Title
    WScript.Quit
  End If
  If Tracks.Count=0 Then Set Tracks=iTunes.LibraryPlaylist.Tracks
  If Tracks.Count=0 Then                ' Can't select ringtones as tracks?
    MsgBox "This script cannot process " & iTunes.BrowserWindow.SelectedPlaylist.Name & ".",vbExclamation,Title
    WScript.Quit
  End If
  ' Check there is a suitable number of suitable tracks to work with
  Count=Tracks.Count
  If Count<Min Or (Count>Max And Max>0) Then
    If Max=0 Then
      MsgBox "Please select " & Min & " or more tracks in iTunes before calling this script!",0,Title
      WScript.Quit
    Else
      MsgBox "Please select between " & Min & " and " & Max & " tracks in iTunes before calling this script!",0,Title
      WScript.Quit
    End If
  End If
  ' Check if the user wants to proceed and how
  Q=Summary
  If Q<>"" Then Q=Q & nl & nl
  If Warn>0 And Count>Warn Then
    Intro=True
    Q=Q & "WARNING!" & nl & "Are you sure you want to process " & GroupDigits(Count) & " tracks"
    If Named Then Q=Q & nl
  Else
    Q=Q & "Process " & GroupDigits(Count) & " track" & Plural(Count,"s "," ")
  End If
  If Named Then Q=Q & "from the " & Source & " playlist"
  Q=Q & "?"
  If Intro Or (Prog And UAC) Then
    If Check Then
      Q=Q & nl & nl 
      Q=Q & "Yes" & tab & ": Process track" & Plural(Count,"s","") & " automatically" & nl
      Q=Q & "No" & tab & ": Preview & confirm each action" & nl
      Q=Q & "Cancel" & tab & ": Abort script"
    End If
    If Kimo Then Q=Q & nl & nl & "NB: Disable ''Keep iTunes Media folder organised'' preference before use."
    If Prog And UAC Then
      Q=Q & nl & nl & "NB: Disable User Access Control to allow progess bar to operate" & nl
      Q=Q & "or change the declaration ''Prog=True'' to ''Prog=False''."
      Prog=False
    End If
    If Check Then
      R=MsgBox(Q,vbYesNoCancel+vbQuestion,Title)
    Else
      R=MsgBox(Q,vbOKCancel+vbQuestion,Title)
    End If
    If R=vbCancel Then WScript.Quit
    If R=vbYes or R=vbOK Then
      Check=False
    Else
      Check=True
    End If
  End If 
  If Check Then Prog=False      ' Suppress progress bar if prompting for user input
End Sub


' Group digits with thousands separator
Function GroupDigits(N)
  GroupDigits=FormatNumber(N,0,,,-1)
End Function


' Return relevant string depending on whether value is plural or singular
' Modified 2011-10-04
Function Plural(V,P,S)
  If V=1 Then Plural=S Else Plural=P
End Function


' Loop through track selection processing suitable items
Sub ProcessTracks
  Dim C,I,N,Q,R,T
  N=0
  If Prog Then                  ' Create ProgessBar
    Set PB=New ProgBar
    PB.SetTitle Title
    PB.Show
  End If
  Clock=0 : StartTimer
  For I=Count To 1 Step -1      ' Work backwards in case edit removes item from selection
    N=N+1                 
    If Prog Then
      PB.SetStatus Status(N)
      PB.Progress N-1,Count
    End If
    Set T=Tracks.Item(I)
    If Prog Then PB.SetInfo Info(T)
    If T.Kind=1 Or T.Kind=3 Then        ' Ignore tracks which can't change
      If Updateable(T) Then     ' Ignore tracks which won't change
        If Check Then           ' Track by track confirmation
          Q=Prompt(T)
          StopTimer             ' Don't time user inputs 
          R=MsgBox(Q,vbYesNoCancel+vbQuestion,Title)
          StartTimer
          Select Case R
          Case vbYes
            C=True
          Case vbNo
            C=False
            S=S+1               ' Increment skipped tracks
          Case Else
            Quit=True
            Exit For
          End Select          
        Else
          C=True
        End If
        If C Then               ' We have a valid track, now do something with it
          Action T
        End If
      Else
        If T.Location<>"" Then V=V+1    ' Increment unchanging tracks, exclude missing ones
      End If
    Else
      MsgBox T.Kind
    End If 
    P=P+1                       ' Increment processed tracks
    If Quit Then Exit For       ' Abort loop on user request
  Next
  StopTimer
  If Prog And Not Quit Then
    PB.Progress Count,Count
    WScript.Sleep 500
    PB.Close
  End If
End Sub


' Output report
Sub Report
  If Not Outro Then Exit Sub
  Dim T
  If Quit Then T="Script aborted!" & nl & nl Else T=""
  T=T & GroupDigits(P) & " track" & Plural(P,"s","")
  If P<Count Then T=T & " of " & count
  T=T & Plural(P," were"," was") & " processed of which " & nl
  If V>0 Then
    T=T & GroupDigits(V) & " did not need updating"
    If (U>0)+(S>0)+(M>0)<-1 Then
      T=T & "," & nl
    ElseIf (U>0)+(S>0)+(M>0)=-1 Then
      T=T & " and" & nl
    End If
  End If
  If U>0 Or V=0 Then
    T=T & GroupDigits(U) & Plural(U," were"," was") & " updated"
    If (S>0)+(M>0)<-1 Then
      T=T & "," & nl
    ElseIf (S>0)+(M>0)=-1 Then
      T=T & " and" & nl
    End If
  End If
  If S>0 Then
    T=T & GroupDigits(S) & Plural(S," were"," was") & " skipped"
    If M>0 Then T=T & " and" & nl
  End If
  If M>0 Then T=T & GroupDigits(M) & Plural(M," were"," was") & " missing"
  T=T & "."
  If Timing Then 
    T=T & nl & nl
    If Check Then T=T & "Processing" Else T=T & "Running"
    T=T & " time: " & FormatTime(Clock)
  End If
  MsgBox T,vbInformation,Title
End Sub


' Start timing event
Sub StartEvent
  T2=Timer
End Sub


' Start timing session
Sub StartTimer
  T1=Timer
End Sub


' Stop timing event and display elapsed time in debug section of Progress Bar
Sub StopEvent
  If Prog Then
    T2=Timer-T2
    If T2<0 Then T2=T2+86400    ' Watch for timer running over midnight
    If Debug Then PB.SetDebug "<br>Last iTunes call took " & FormatTime(T2) 
  End If  
End Sub


' Stop timing session and add elapased time to running clock
Sub StopTimer
  Clock=Clock+Timer-T1
  If Clock<0 Then Clock=Clock+86400     ' Watch for timer running over midnight
End Sub


' Detect if User Access Control is enabled, UAC prevents use of progress bar
Function UAC
  Const HKEY_LOCAL_MACHINE=&H80000002
  Const KeyPath="Software\Microsoft\Windows\CurrentVersion\Policies\System"
  Const KeyName="EnableLUA"
  Dim Reg,Value
  Set Reg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 	  ' Use . for local computer, otherwise could be computer name or IP address
  Reg.GetDWORDValue HKEY_LOCAL_MACHINE,KeyPath,KeyName,Value	  ' Get current property
  If IsNull(Value) Then UAC=False Else UAC=(Value<>0)
End Function


' ==================
' Progress Bar Class
' ==================

' Progress/activity bar for vbScript implemented via IE automation
' Can optionally rebuild itself if closed or abort the calling script
Class ProgBar
  Public Cells,Height,Width,Respawn,Title,Version
  Private Active,Blank,Dbg,Filled(),FSO,IE,Info,NextOn,NextOff,Status,SHeight,SWidth,Temp

' User has closed progress bar, abort or respwan?
' Modified 2011-10-09
  Public Sub Cancel()
    If Respawn And Active Then
      Active=False
      If Respawn=1 Then
        Show                    ' Ignore user's attempt to close and respawn
      Else
        Dim R
        StopTimer               ' Don't time user inputs 
        R=MsgBox("Abort Script?",vbExclamation+vbYesNoCancel,Title)
        StartTimer
        If R=vbYes Then
          On Error Resume Next
          CleanUp
          Respawn=False
          Quit=True             ' Global flag allows main program to complete current task before exiting
        Else
          Show                  ' Recreate box if closed
        End If  
      End If        
    End If
  End Sub

' Delete temporary html file  
  Private Sub CleanUp()
    FSO.DeleteFile Temp         ' Delete temporary file
  End Sub
  
' Close progress bar and tidy up
' Modified 2011-10-04
  Public Sub Close()
    On Error Resume Next        ' Ignore errors caused by closed object
    If Active Then
      Active=False              ' Ignores second call as IE object is destroyed
      IE.Quit                   ' Remove the progess bar
      CleanUp
    End If    
 End Sub
 
' Initialize object properties
  Private Sub Class_Initialize()
    Dim I,Items,strComputer,WMI
    ' Get width & height of screen for centering ProgressBar
    strComputer="."
    Set WMI=GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set Items=WMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
    'Get the OS version number (first two)
    For Each I in Items
      Version=Left(I.Version,3)
    Next
    Set Items=WMI.ExecQuery ("Select * From Win32_DisplayConfiguration")
    For Each I in Items
      SHeight=I.PelsHeight
      SWidth=I.PelsWidth
    Next
    If Debug Then
      Height=140                ' Height of containing div
    Else
      Height=100                ' Reduce height if no debug area
    End If
    Width=300                   ' Width of containing div
    Respawn=True                ' ProgressBar will attempt to resurect if closed
    Blank=String(50,160)        ' Blanks out "Internet Explorer" from title
    Cells=25                    ' No. of units in ProgressBar, resize window if using more cells
    ReDim Filled(Cells)         ' Array holds current state of each cell
    For I=0 To Cells-1
      Filled(I)=False
    Next
    NextOn=0                    ' Next cell to be filled if busy cycling
    NextOff=Cells-5             ' Next cell to be cleared if busy cycling
    Dbg="&nbsp;"                ' Initital value for debug text
    Info="&nbsp;"               ' Initital value for info text
    Status="&nbsp;"             ' Initital value for status text
    Title="Progress Bar"        ' Initital value for title text
    Set FSO=CreateObject("Scripting.FileSystemObject")          ' File System Object
    Temp=FSO.GetSpecialFolder(2) & "\ProgBar.htm"               ' Path to Temp file
  End Sub

' Tidy up if progress bar object is destroyed
  Private Sub Class_Terminate()
    Close
  End Sub
 
' Display the bar filled in proportion X of Y
  Public Sub Progress(X,Y)
    Dim F,I,L,S,Z
    If X<0 Or X>Y Or Y<=0 Then
      MsgBox "Invalid call to ProgessBar.Progress, variables out of range!",vbExclamation,Title
      Exit Sub
    End If
    Z=Int(X/Y*(Cells))
    If Z=NextOn Then Exit Sub
    If Z=NextOn+1 Then
      Step False
    Else
      If Z>NextOn Then
        F=0 : L=Cells-1 : S=1
      Else
        F=Cells-1 : L=0 : S=-1
      End If
      For I=F To L Step S
        If I>=Z Then
          SetCell I,False
        Else
          SetCell I,True
        End If
      Next
      NextOn=Z
    End If
  End Sub

' Clear progress bar ready for reuse  
' Modified 2011-10-16
  Public Sub Reset
    Dim C
    For C=Cells-1 To 0 Step -1
      IE.Document.All.Item("P",C).classname="empty"
      Filled(C)=False
    Next
    NextOn=0
    NextOff=Cells-5   
  End Sub
  
' Directly set or clear a cell
  Public Sub SetCell(C,F)
    On Error Resume Next        ' Ignore errors caused by closed object
    If F And Not Filled(C) Then
      Filled(C)=True
      IE.Document.All.Item("P",C).classname="filled"
    ElseIf Not F And Filled(C) Then
      Filled(C)=False
      IE.Document.All.Item("P",C).classname="empty"
    End If
  End Sub 
 
' Set text in the Dbg area
  Public Sub SetDebug(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Dbg=T
    IE.Document.GetElementById("Debug").InnerHTML=T
  End Sub

' Set text in the info area
  Public Sub SetInfo(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Info=T
    IE.Document.GetElementById("Info").InnerHTML=T
  End Sub

' Set text in the status area
  Public Sub SetStatus(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Status=T
    IE.Document.GetElementById("Status").InnerHTML=T
  End Sub

' Set title text
  Public Sub SetTitle(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Title=T
    IE.Document.Title=T & Blank
  End Sub
  
' Create and display the progress bar  
  Public Sub Show()
    Const HKEY_CURRENT_USER=&H80000001
    Const KeyPath="Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN"
    Const KeyName="iexplore.exe"
    Dim File,I,Reg,State,Value
    Set Reg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 	' Use . for local computer, otherwise could be computer name or IP address
    'On Error Resume Next        ' Ignore possible errors
    ' Make sure IE is set to allow local content, at least while we get the Progress Bar displayed
    Reg.GetDWORDValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	' Get current property
    State=Value	 							  ' Preserve current option
    Value=0		    							' Set new option 
    Reg.SetDWORDValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	' Update property
    'If Version<>"5.1" Then Prog=False : Exit Sub      ' Need to test for Vista/Windows 7 with UAC
    Set IE=WScript.CreateObject("InternetExplorer.Application","Event_")
    Set File=FSO.CreateTextFile(Temp, True)
    With File
      .WriteLine "<!doctype html>"
      .WriteLine "<html><head><title>" & Title & Blank & "</title>"
      .WriteLine "<style type='text/css'>"
      .WriteLine ".border {border: 5px solid #DBD7C7;}"
      .WriteLine ".debug {font-family: Tahoma; font-size: 8.5pt;}"
      .WriteLine ".empty {border: 2px solid #FFFFFF; background-color: #FFFFFF;}"
      .WriteLine ".filled {border: 2px solid #FFFFFF; background-color: #00FF00;}"
      .WriteLine ".info {font-family: Tahoma; font-size: 8.5pt;}"
      .WriteLine ".status {font-family: Tahoma; font-size: 10pt;}"
      .WriteLine "</style>"
      .WriteLine "</head>"
      .WriteLine "<body scroll='no' style='background-color: #EBE7D7'>"
      .WriteLine "<div style='display:block; height:" & Height & "px; width:" & Width & "px; overflow:hidden;'>"
      .WriteLine "<table border-width='0' cellpadding='2' width='" & Width & "px'><tr>"
      .WriteLine "<td id='Status' class='status'>" & Status & "</td></tr></table>"
      .WriteLine "<table class='border' cellpadding='0' cellspacing='0' width='" & Width & "px'><tr>"
      ' Write out cells
      For I=0 To Cells-1
	      If Filled(I) Then
          .WriteLine "<td id='p' class='filled'>&nbsp;</td>"
        Else
          .WriteLine "<td id='p' class='empty'>&nbsp;</td>"
        End If
      Next
	    .WriteLine "</tr></table>"
      .WriteLine "<table border-width='0' cellpadding='2' width='" & Width & "px'><tr><td>"
      .WriteLine "<span id='Info' class='info'>" & Info & "</span><br>"
      .WriteLine "<span id='Debug' class='debug'>" & Dbg & "</span></td></tr></table>"
      .WriteLine "</div></body></html>"
    End With
    ' Create IE automation object with generated HTML
    With IE
      .width=Width+30           ' Increase if using more cells
      .height=Height+55         ' Increase to allow more info/debug text
      If Version>"5.1" Then     ' Allow for bigger border in Vista/Widows 7
        .width=.width+10
        .height=.height+10
      End If        
      .left=(SWidth-.width)/2
      .top=(SHeight-.height)/2
      .navigate "file://" & Temp
      '.navigate "http://samsoft.org.uk/progbar.htm"
      .addressbar=False
      .menubar=False
      .resizable=False
      .toolbar=False
      On Error Resume Next      
      .statusbar=False          ' Causes error on Windows 7 or IE 9
      On Error Goto 0
      .visible=True             ' Causes error if UAC is active
    End With
    Active=True
    ' Restore the user's property settings for the registry key
    Value=State		    					' Restore option
    Reg.SetDWORDValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	  ' Update property 
    Exit Sub
  End Sub
 
' Increment progress bar, optionally clearing a previous cell if working as an activity bar
' Modified 2011-10-05
  Public Sub Step(Clear)
    SetCell NextOn,True : NextOn=(NextOn+1) Mod Cells
    If Clear Then SetCell NextOff,False : NextOff=(NextOff+1) Mod Cells
  End Sub

' Self-timed shutdown
' Modified 2011-10-05 
  Public Sub TimeOut(S)
    Dim I
    Respawn=False                ' Allow uninteruppted exit during countdown
    For I=S To 2 Step -1
      SetDebug "<br>Closing in " & I & " seconds" & String(I,".")
      WScript.sleep 1000
    Next
      SetDebug "<br>Closing in 1 second."
      WScript.sleep 1000
    Close
  End Sub 
    
End Class


' Fires if progress bar window is closed, can't seem to wrap up the handler in the class
' Modified 2011-10-04
Sub Event_OnQuit()
  PB.Cancel
End Sub


' ==============
' End of listing
' ==============