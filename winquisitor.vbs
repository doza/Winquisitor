'*******************************************************************************
' Script Name:  winquisitor.vbs
' Author:       Mike Cardosa
' Website:      http://www.winquisitor.org
' Last Updated: September 22, 2010 
'
' Description:  Performs tests against one or more Windows computers and outputs
'               the results in the given format.
'               Possible tests:
'                   - Existence of a Windows Service
'                   - Existence and version of a file
'                   - Existence of an MS patch
'                   - Existence of a Registry key
'                   - Existence of a running process
'                   - Contents of a Registry value
'                   - Existence of a local user account
'                   - List members of a local group
'                   - Execute a custom WMI query
'               Possible output formats:
'                   - Human readable (on StdOut)
'                   - XML
'                   - CSV
' Disclaimer:   The author makes no representations about the suitability
'               of this software for any purpose.  This software is provided
'               AS IS and without any express or implied warranties, 
'               including, without limitation, the implied warranties of 
'               merchantability and fitness for a particular purpose. The
'               entire risk arising out of the use or performance of this script 
'               and documentation remains with you. In no event shall the author,
'               or anyone else involved in the creation, production, or delivery
'               of the scripts be liable for any damages whatsoever (including,
'               without limitation, damages for loss of business profits, 
'               business interruption, loss of business information, or other
'               pecuniary loss) arising out of the use of or inability to use
'               the script or documentation, even if the author has been
'               advised of the possibility of such damages.
'
'*******************************************************************************

Option Explicit

'*******************************************************************************
' Constants
'*******************************************************************************
Const Version      = "0.1.5"


'-------------------------------------------------------------------------------
' Constants for connecting to remote machines
'-------------------------------------------------------------------------------
Const WbemAuthenticationLevelPktPrivacy = 6


'-------------------------------------------------------------------------------
' Constants for opening text files
'-------------------------------------------------------------------------------
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8


'-------------------------------------------------------------------------------
' Constants for interacting with the Registry
'-------------------------------------------------------------------------------
Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA       = &H80000006

Const REG_SZ              = 1
Const REG_EXPAND_SZ       = 2
Const REG_BINARY          = 3
Const REG_DWORD           = 4
Const REG_MULTI_SZ        = 7


'-------------------------------------------------------------------------------
' Constants for connecting to the WinNT provider
'-------------------------------------------------------------------------------
Const ADS_SECURE_AUTHENTICATION = &H1
Const ADS_USE_ENCRYPTION        = &H2


'-------------------------------------------------------------------------------
' Constants for handling records in output files
'-------------------------------------------------------------------------------
Const RECORD_SEPARATOR     = "//"
Const PROPERTY_SEPARATOR   = " || "


'*******************************************************************************
' Variables
'*******************************************************************************

'-------------------------------------------------------------------------------
' Variables for controlling the script output
'-------------------------------------------------------------------------------
Dim bQuiet             : bQuiet = False
Dim bVerbose           : bVerbose = False
Dim bVeryVerbose       : bVeryVerbose = False
Dim bDebug             : bDebug = False
Dim bHaveVerboseLevel  : bHaveVerboseLevel = False


'-------------------------------------------------------------------------------
' Variables for controlling output to a specified file
'-------------------------------------------------------------------------------
Dim sOutputFile        : sOutputFile = ""
Dim bOutputFile        : bOutputFile = False
Dim sOutputFormat      : sOutputFormat = "XML"
Dim iOpenOutputFile    : iOpenOutputFile = ForWriting
Dim oOutputFile
Dim oOutputFileFSO
Dim bAppendToXML       : bAppendToXML = False
Dim bAppend            : bAppend = False


'-------------------------------------------------------------------------------
' Variables for tracking and running tests
'-------------------------------------------------------------------------------
Dim iTestCount         : iTestCount = 0      ' Tracks the number of tests we will run
Dim iNumTests          : iNumTests = -1      ' Keeps track of the number of tests we have
Dim iCurrentTest  
Dim sCurrentTestType                         ' 
Dim sCurrentTestValue

Dim bPing              : bPing = True        ' Boolean for indicating whether or not we should ping each target machine before testing
Dim sPingStatus

Dim bRunTests          : bRunTests = False
Dim bNeedRegObj        : bNeedRegObj = False ' Boolean to let us know if we need to connect to the Registry on the target
Dim bHaveRegObj        : bHaveRegObj = False

Dim bResultDetails                           ' If true, provide the full query results of a custom query
bResultDetails = False                       '   If false, only provide a count of the results

Dim aTests(254,2)

Dim aTargets()                               ' Array for storing the target machines
Dim sCurrentTarget                           ' The name of the current target machine we are working with

Dim oMember

Dim aLocalGroups()                           ' Array for storing local groups to enumerate
Dim sLocalGroup                              ' The current group we are testing

Dim sTestResult                              ' Stores the result of the current test

Dim bCScript           : bCScript = False    ' Boolean for tracking whether the script was called with CScript

Dim bWebXSL            : bWebXSL = False     ' Boolean for tracking whether or not to include the XSL file from winquisitor.org
Dim sXSLFile           : sXSLFile = ""       ' Name of the XSL file to include
Dim sScanInfo


Dim sUsername          : sUsername = ""      ' Username to connect with if not current user
Dim sPassword          : sPassword = ""      ' Password to use for the connection if a username is specified


'-------------------------------------------------------------------------------
' Common objects for the script
'-------------------------------------------------------------------------------
Dim oWshShell          : Set oWshShell = WScript.CreateObject("WScript.Shell")
Dim oWbemLocator       : Set oWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Dim oStdOut            : Set oStdOut = WScript.StdOut
Dim oStdErr            : Set oStdErr = WScript.StdErr
Dim oStdIn             : Set oStdIn  = WScript.StdIn
Dim oWMIService                              ' Object used to connect to each target machine
Dim oDefault                                 ' Object used to connect to the default namespace
Dim oReg                                     ' Object used for interacting with the Registry
Dim oGroup                                   ' Object used for enumerating local groups
Dim oNS                : Set oNS = GetObject("WinNT:")
Dim sNameSpace                               ' The WMI namespace we are connecting to
Dim iTimer


'-------------------------------------------------------------------------------
' Script usage displayed when displayUsage() is called
'-------------------------------------------------------------------------------                  
Dim sUsage
sUsage = vbCRLF & WScript.ScriptName & " v" & Version & " ( http://winquisitor.org ) " & vbCRLF & _
                  vbCRLF & _
                  "USAGE:" & vbCRLF & _
                  "=====================" & vbCRLF & _
                  vbCRLF & _
                  "cscript [ //nologo ] " & WScript.ScriptName & " [ -h|--help ]" & vbCRLF & _
                  vbCRLF & _
                  "cscript [ //nologo ] " & WScript.ScriptName & " { test(s) } [ output ] { target specification }" & vbCRLF & _
                  vbCRLF & _
                  vbCRLF & _
                  "PARAMETERS:" & vbCRLF & _
                  "=====================" & vbCRLF & _
                  vbCRLF & _
                  " OUTPUT:" & vbCRLF & _
                  " --------------------" & vbCRLF & _
                  "  -h,--help                    Display this usage screen" & vbCRLF & _
                  "  -v                           Enable verbose output" & vbCRLF & _
                  "  -vv                          Enable very verbose output" & vbCRLF & _
                  "  -d,--debug                   Enable debugging output" & vbCRLF & _
                  "  -q,--quiet                   Suppress output" & vbCRLF & _
                  "  -oC:file                     Output CSV results to the given file" & vbCRLF & _
                  "  -oX:file                     Output XML results to the given file" & vbCRLF & _
                  "  -xsl:file                    Reference the given XSL document in the" & vbCRLF & _
                  "                                 XML output file instead of the default" & vbCRLF & _
                  "                                 winquisitor.xsl" & vbCRLF & _
                  "  --web-xsl                    Reference the XSL file hosted on winquisitor.org" & vbCRLF & _
                  "                                 in the XML output file instead of the" & vbCRLF & _
                  "                                 default winquisitor.xsl" & vbCRLF & _
                  "                                 Note: This will not work in Firefox because" & vbCRLF & _
                  "                                 FF will not parse XSL files from a different" & vbCRLF & _
                  "                                 scope than the XML file." & vbCRLF & _
                  "  --append-output              Append to the given output file instead of " & vbCRLF & _
                  "                                 overwriting" & vbCRLF & _
                  vbCRLF & _
                  " TARGET SPECIFICATION:" & vbCRLF & _
                  " --------------------" & vbCRLF & _
                  "  -t,--target:computer         Add the given computer to the list of computers" & vbCRLF & _
                  "                                 to test" & vbCRLF & _
                  "  -T,--target-file:file        Read targets from the given file" & vbCRLF & _
                  "                                 (one target per line)" & vbCRLF & _
                  "  -np,--no-ping                Do not ping targets before trying to connect" & vbCRLF & _
                  "  -u,--username:username       Connect to targets with the given username" & vbCRLF & _
                  "  -p,--password:password       Connect to targets with the given password" & vbCRLF & _
                  "                                 If a username was given and a password was" & vbCRLF & _
                  "                                 not specified, then the user will be prompted" & vbCRLF & _
                  "                                 for a password." & vbCRLF & _
                  vbCRLF & _
                  " TESTS:" & vbCRLF & _
                  " --------------------" & vbCRLF & _
                  "  -f,--file:file               Test the existence and version of the given file" & vbCRLF & _
                  "  -s,--service:service         Test the state of the given service" & vbCRLF & _
                  "  -pa,--patch:patch            Test whether a given patch has been applied" & vbCRLF & _
                  "  -pr,--process:process        Test whether or not a process is running" & vbCRLF & _
                  "  -rk,--registry-key:key       Test the existence and/or value of the" & vbCRLF & _
                  "                                 given registry key" & vbCRLF & _
                  "  -rv,--regisry-value:value    Test the given registry value" & vbCRLF & _
                  "  -lu,--local-user:username    Test the existence of the given user" & vbCRLF & _
                  "  -lg,--local-group:groupname  Enumerate the members of the given local group" & vbCRLF & _
                  "  -cq,--custom-query:query     WMI query against the CIMV2 namespace" & vbCRLF & _
                  "  --result-detail              Provide detailed results instead of a summary." & vbCRLF & _
                  "                                 Any properties and values will be enumerated." & vbCRLF & _
                  vbCRLF & _
                  vbCRLF & _
                  "EXAMPLES:" & vbCRLF & _
                  "=====================" & vbCRLF & _
                  vbCRLF & _
                  " EXAMPLE 1:" & vbCRLF & _
                  " --------------------" & vbCRLF & _
                  "  Test for the Alerter service on machines 192.168.1.10 and 192.168.1.11" & vbCRLF & _
                  "    and record results in XML format to results.xml" & vbCRLF & _
                  vbCRLF & _
                  "   " & WScript.ScriptName & " -t:192.168.1.10 -t:192.168.1.11 -s:Alerter -oX:results.xml" & vbCRLF & _
                  vbCRLF & _
                  vbCRLF & _
                  " EXAMPLE 2:" & vbCRLF & _
                  " --------------------" & vbCRLF & _
                  "  Test for the existence of the file ""C:\Windows\system32\evil.exe"" and" & vbCRLF & _
                  "    the running process trojan.exe against 192.168.1.10, 192.168.1.1, and all" & vbCRLF & _
                  "    hosts listed in targets.txt. Record detailed results in XML format" & vbCRLF & _
                  "    to results.xml" & vbCRLF & _
                  vbCRLF & _
                  "   " & WScript.ScriptName & " -t:192.168.1.10 -t:192.168.1.11 -T:targets.txt" & vbCRLF & _
                  "     -f:""C:\Windows\system32\evil.exe"" -p:""trojan.exe"" -oX:results.xml" & vbCRLF & _
                  "     --result-detail" & vbCRLF & _                 
                  vbCRLF & _
                  vbCRLF & _
                  " EXAMPLE 3:" & vbCRLF & _
                  " --------------------" & vbCRLF & _
                  "  Check for patch KB890046 and run a custom query against 192.168.1.11" & vbCRLF & _
                  "    displaying detailed results. Do not ping the target first. Append the" & vbCRLF & _
                  "    results in CSV format to results.csv" & vbCRLF & _
                  vbCRLF & _
                  "   " & WScript.ScriptName & " -t:192.168.1.11 -np -pa:KB890046 -oC:results.csv" & vbCRLF & _
                  "     -cq:""select caption from win32_useraccount"" --result-detail --append-output" & vbCRLF & _
                  vbCRLF   
                  
                                                     
'*******************************************************************************
' Main script
'*******************************************************************************

iTimer = Timer

bCScript = checkCScript()          ' Confirm that the script was run using CScript
If Not bCScript Then
	WScript.Echo WScript.ScriptName & " must be run using CScript"
	WScript.Quit	
End If

debug("Processing command line arguments")
Call processCommandLine()       ' Go through the arguments and flags provided on the command line

If Not bDebug Then On Error Resume Next

debug("Finished processing command line")
' If we have targets, remove any duplicates
If getUBound(aTargets) > -1 Then
  ' Remove any duplicate target machines
  debug("Removing duplicate targets")
  Call removeArrayDuplicates(aTargets)
Else
	' No targets were given to the script. Quit
	fatalError("No targets specified")
End If

' Check if we have tests to run
debug("Checking if we have tests")
If iNumTests < 0 Then
	' If we have no tests then we exit with an error
	' Why would we run without any tests?
	fatalError("No tests specified")
End If
debug("We do have tests")

' If an output file was given, we need to prepare it
If bOutputFile Then
	debug(formatOutput("Preparing output file:",sOutputFile))
	Set oOutputFileFSO = CreateObject("Scripting.FileSystemObject")

  ' If the file exists
	If oOutputFileFSO.FileExists(sOutputFile) Then
		debug(sOutputFile & " already exists")
		If iOpenOutputFile = ForWriting Then
			debug("Will overwrite file")
		Else
			debug("Will append to file")
			' Need to clean up an XML file if we are going to append to it
			If sOutputFormat = "XML" Then
			  Dim aLines
			  Dim oInFile
			  Dim sFileContents
			  Dim i
			  
			  debug("Need to remove the trailing element to append to an XML file")
			  debug("Reading current file")
			  Set oInFile = oOutputFileFSO.OpenTextFile(sOutputFile, ForReading)
		    sFileContents = oInFile.ReadAll
			  oInFile.close
			  
			  aLines = split(sFileContents,vbCRLF)
			  
			  Set oOutputFile = oOutputFileFSO.OpenTextFile(sOutputFile, ForWriting)
			  For i = 0 To getUBound(aLines)
			    Select Case Trim(LCase(aLines(i)))
			      Case "</winquisitor_audit>"
			        bAppendToXML = True
			      Case ""
			      Case Else
			      	oOutputFile.WriteLine aLines(i)
			    End Select
			  Next 
			  oOutputFile.Close
		  End If
		End If
	  Set oOutputFile = oOutputFileFSO.OpenTextFile(sOutputFile, iOpenOutputFile)
	Else
		' The file does not exist so we need to create it
		debug("Creating output file")
		Set oOutputFile = oOutputFileFSO.CreateTextFile(sOutputFile)
		If Err <> 0 Then
			fatalError("Could not create output file " & sOutputFile & "-" & Err.Description)
		End If
	End If
	
  If Err <> 0 Then
  	fatalError("Error with output file: " & sOutputFile & " - " & Err.Description)
  End If
  
  Call prepareOutputFile()
  debug("Output file prepared successfully")
End If


' If we were given a username without a password, then we need to ask for the
'   password to use
If sUsername <> "" And sPassword = "" Then
	Dim oPassword 
	debug("Need to prompt for password")
	' You can only use ScriptPW.Password on XP and 2003.
  ' If the ScriptPW.Password object is not available, the call to CreateObject
  '  will fail and take the script with it. Need to ignore that error and
  '  provide another way to accept the password.
  On Error Resume Next
	  Set oPassword = WScript.CreateObject("ScriptPW.Password")
	
	  If Err <> 0 Then
	  	debug("Could not create ScriptPW.Password. Will need to use StdIn")
	  	'fatalError("Cannot prompt for password prior to Windows XP" & vbCRLF & Err.Description)
	  	oStdOut.WriteLine vbNewLine & "WARNING: scriptpw.dll does not appear to be registered."
	  	oStdOut.WriteLine vbNewLine & "         Your password will not be masked!"
	  	oStdOut.Write vbNewLine & "Enter password for " & sUsername & " or press CTRL+C to exit: "
	  	sPassword = oStdIn.ReadLine
	  	debug("Password received using StdIn")
	  Else
	  	debug("Requesting password using ScriptPW.Password")
	  	oStdOut.WriteLine vbNewLine & "Enter password for user " & sUsername & ":"
	  	sPassword = oPassword.GetPassword()
	  	debug("Password received using ScriptPW.Password")
	  End If
	On Error Goto 0
End If

' Loop through the aTargets and run specified tests
'  against each one
For each sCurrentTarget in aTargets
  ' Turn on error handling
  If Not bDebug Then On Error Resume Next
  ' Clear any previous errors from the buffer
  Err.Clear
  output("Processing: " & sCurrentTarget)

  If bOutputFile Then
  	Call recordNewTarget(sCurrentTarget)
  End If
  
  bRunTests = True
  If bPing Then
  	veryVerbose("Pinging " & sCurrentTarget)
  	sPingStatus = pingStatus(sCurrentTarget)
  	If sPingStatus = "Success" Then
  	  output("   Ping succeeded")
    Else 
    	displayError("   Ping failed: " & sPingStatus)
    	recordConnectFailure("Ping failed: " & sPingStatus)
      bRunTests = False
    End If
  End If
  
  If bRunTests Then
  
    If sUsername <> "" Then  
  	  sNamespace = "root\cimv2"
  	  
  	  debug("Connecting to WMI with alternate credentials")
		  Set oWMIService = oWbemLocator.ConnectServer _
		      (sCurrentTarget, sNamespace, sUsername, sPassword)
		  oWMIService.Security_.authenticationLevel = WbemAuthenticationLevelPktPrivacy
		  'oWMIService.Security_.ImpersonationLevel = 3
	  Else	
	  	debug("Connecting with current credentials")
      Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
                                  sCurrentTarget & "\root\cimv2")
    End If
  
      If IsObject(oWMIService) And Err = 0 Then
      	recordConnectSuccess()
      	
      	If bNeedRegObj Then
      		If sUsername <> "" Then
      			debug("Connecting to registry with alternate credentials")
      		  sNameSpace = "root\default"
      		  
      		  Set oDefault = oWbemLocator.ConnectServer _
		                          (sCurrentTarget, sNamespace, sUsername, sPassword)
		        oDefault.Security_.authenticationLevel = WbemAuthenticationLevelPktPrivacy
		        Set oReg = oDefault.Get("StdRegProv")
      	  Else
      	  	debug("Connecting to registry with current credentials")
      		  Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
                                 sCurrentTarget & "\root\default:StdRegProv")
          End If
          
          If IsObject(oReg) And Err = 0 Then
            bHaveRegObj = True
            debug("Successfully created Registry object")
          Else
          	bHaveRegObj = False
          	displayError("Could not query registry on " & sCurrentTarget)
          End If
      	End If
      	
      	' Run all the tests
      	For iCurrentTest = 0 To iNumTests
      	  sCurrentTestType = aTests(iCurrentTest,0)
      	  sCurrentTestValue = aTests(iCurrentTest,1)
      	  sTestResult = ""
      	  
      	  Select Case sCurrentTestType
      	  	Case "Service"
      	  	  sTestResult = checkService(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"Service",sCurrentTestValue,sTestResult)
      	  	Case "File"
      	  	  sTestResult = checkFile(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"File",sCurrentTestValue,sTestResult)
      	  	Case "Patch"
      	  	  sTestResult = checkPatch(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"Patch",sCurrentTestValue,sTestResult)      	  	  
      	  	Case "Process"
      	  	  sTestResult = checkProcess(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"Process",sCurrentTestValue,sTestResult)
      	  	Case "Registry Key"
     	  	  	sTestResult = checkRegistryKey(sCurrentTestValue)
     	  	  	Call recordResult(sCurrentTarget,"Registry Key",sCurrentTestValue,sTestResult)
      	  	Case "Registry Value"
     	  	  	sTestResult = checkRegistryValue(sCurrentTestValue)
     	  	  	Call recordResult(sCurrentTarget,"Registry Value",sCurrentTestValue,sTestResult)
      	  	Case "Local User"
      	  	  sTestResult = checkLocalUser(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"Local User",sCurrentTestValue,sTestResult)
      	  	Case "Local Group"
      	  	  sTestResult = enumLocalGroup(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"Local Group",sCurrentTestValue,sTestResult)
      	  	Case "Custom Query"
      	  	  sTestResult = executeQuery(sCurrentTestValue)
      	  	  Call recordResult(sCurrentTarget,"Custom Query",sCurrentTestValue,sTestResult)      	  	  
      	  	  
      	  End Select
      	
        Next

      Else
  	    ' We couldn't connect to the computer
  	    displayError("   Error: Could not connect to " & sCurrentTarget & " - " & Err.Description)
  	    recordConnectFailure("Could not connect - " & Err.Description)
  	    Err.Clear
  	    
      End If
    'End If

    Set oWMIService = Nothing
    Set oReg = Nothing
  End If
  
  If bOutputFile Then
    Call endCurrentTarget(sCurrentTarget)
  End If
  
  On Error GoTo 0
Next

If bOutputFile And (sOutputFormat = "XML") Then
	Call fullXMLElement("end_date",Date)
	Call fullXMLElement("end_time",Time)
	Call closeXMLElement("scan")
	Call closeXMLELement("winquisitor_audit")
End If

debug("Total run time: " & Round(Timer - iTimer, 4) & " seconds")
  

'*************************************************
'
'         ******  Begin Functions ******
'
'*************************************************

'*************************************************
'
' Function: checkCScript
'
' Description: Checks if the Windows Script Host
'   is CScript.exe
'*************************************************
Function checkCScript()
  Dim sEngine

  sEngine = UCase(Right(WScript.FullName,12))

  If sEngine <> "\CSCRIPT.EXE" Then
	  checkCScript = False
	Else
		checkCScript = True
  End If
    
End Function
'*************************************************
' End Function: checkCScript
'*************************************************


'*************************************************
'
' Sub: processCommandLine
'
' Description: Goes through all of the flags and 
'  arguments provided to the script and builds the
'  tests to run, target machine list, output format,
'  etc. If there were no arguments to the script,
'  display the usage and quit.
'*************************************************
Sub processCommandLine()
	Dim iArguments			'the number of arguments passed to the script
	Dim sFullArgument		'the full string of the current argument we are working with (flag and value)
	Dim aArgument       'an array to hold the flag and value for the argument
	Dim sArgument				'the string of the current argument (no value)
	Dim sArgumentValue	'the string of the value for the current argument (no flag)
	Dim iArgument				'index used for looping through the arguments
		
	iArguments = WScript.Arguments.Count - 1
	
	' If no arguments were passed to the script, display the correct usage and exit
	If iArguments < 0 Then
		displayUsage
	End If
	
	' First we want to check for the verbosity level if it is present on the command line
	' We do this first in case verbosity is set high enough for us to display what
	'   we are doing with the command line arguments. If we don't check it first, we might
	'   miss some output if the verbosity level is specified late in the command line.
	For iArgument = 0 To iArguments
	  If Not bHaveVerboseLevel Then
	    sFullArgument = WScript.Arguments(iArgument)
	    'If InStr(sFullArgument,"-v") Or InStr(sFullArgument,"/v") Or InStr(sFullArgument,"/q") Or InStr(sFullArgument,"-q") Then
	      'sFullArgument = Mid(sFullArgument,2)
	      Select Case sFullArgument
	      	Case "-v","/v"
	      	  bQuiet = False
	      	  bVerbose = True
	      	  bVeryVerbose = False
	      	  bDebug = False
	      	  bHaveVerboseLevel = True
	      	Case "-vv","/vv"
	      	  bQuiet = False
	      	  bVerbose = True
	      	  bVeryVerbose = True
	      	  bDebug = False
	      	  bHaveVerboseLevel = True
	      	Case "-d","/d","/debug","--debug"
	      	  bQuiet = False
	      	  bVerbose = True
	      	  bVeryVerbose = True
	      	  bDebug = True
	      	  bHaveVerboseLevel = True
	      	Case "-q","/q","--quiet","/quiet"
	      	  bQuiet = True
  	    	  bVerbose = False
  	    	  bVeryVerbose = False
  	    	  bDebug = False
  	    	  bHaveVerboseLevel = True
  	    End Select
  	  'End If
  	End If
  Next
	
	' Loop through all the arguments on the command line	
	For iArgument = 0 To iArguments
		sFullArgument = WScript.Arguments(iArgument)

		If ((Left(sFullArgument,1) = "-") Or (Left(sFullArgument,1) = "/")) And (Len(sFullArgument) > 1) Then
			
			' If there is a : in the argument, then we need to separate the argument and value
			If InStr(sFullArgument,":") Then
			  ' Create an array out of the individual argument
			  aArgument = Split(sFullArgument,":",2)
			  ' The actual argument is in the first element of the array
			  sArgument = aArgument(0)
			  ' the value for the argument (if there is one) is the second element
			  If Trim(aArgument(1)) <> "" Then
			    sArgumentValue = Trim(aArgument(1))
			  Else
			  	sArgumentValue = ""
			  End If
			Else
				' The argument has no : so it might be a flag
				sArgument = sFullArgument
				sArgumentValue = ""
			End If
		  
		  Select Case sArgument
		  
		  '----------------------------------------
		  ' Script arguments
		  '----------------------------------------
		  
		  ' Look for help flag
		  Case "-h","-?","--help","/h","/help","/?"
		    displayUsage
		    WScript.Quit
		    
		  ' Look for any output variables
		  Case "-v","-vv","-d","-q","/v","/vv","/d","/q","--debug","/debug","--quiet","/quiet"
		    ' Nothing to do since we already checked verbosity level earlier
		  
		  
		  '------------------------
		  ' Target options
		  '------------------------
		  Case "-t","--target","/t","/target"
		    If sArgumentValue <> "" Then
		      ' If a value was given with the argument, add it to the array
		      veryVerbose(formatOutput("Adding target:",sArgumentValue))
		      Call addArrayItem(aTargets,sArgumentValue)
		    Else
		    	' No value was specified
		    	fatalError("No value given for: " & sArgument)
		    End If
		  Case "-T","--targetfile","/T","/Targetfile"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Reading target file:",sArgumentValue))
		      Call readTargetFile(sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If
		  Case "-np","--no-ping","/np","/no-ping"
		    veryVerbose("Will not ping targets")
		    bPing = False


		  '------------------------
		  ' Alternate credentials
		  '------------------------
		  Case "-u","--username","/u","/username"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Username specified:",sArgumentValue))
		      sUsername = sArgumentValue
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If
		  Case "-p","--password","/p","/password"
		    If sArgumentValue <> "" Then
		    	veryVerbose("Password specified")
		      sPassword = sArgumentValue
		    Else
		    	veryVerbose("No password specified.")
		    	
		    End If		    
		    
		  '------------------------
		  ' Test arguments
		  '------------------------
		  Case "-f","--file","/f","/file"
		    If sArgumentValue <> "" Then
		    	veryVerbose(formatOutput("Adding file test:",sArgumentValue))
          Call addTest("File",sArgumentValue)
		    Else
		    	' No value was specified
		    	fatalError("No value given for: " & sArgument)
		    End If

		  Case "-s","--service","/s","/service"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding service:",sArgumentValue))
		      Call addTest("Service",sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If

		  Case "-pa","--patch","/pa","/patch"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding patch:",sArgumentValue))
		      Call addTest("Patch",sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If

		  Case "-pr","--process","/pr","/process"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding process:",sArgumentValue))
		      Call addTest("Process",sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If
		    
		  Case "-rk","--registry-key","/rk","/registry-key"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding registry key:",sArgumentValue))
		      Call addTest("Registry Key",sArgumentValue)
		      bNeedRegObj = True
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If
		    
		  Case "-rv","--registry-value","/rv","/registry-value"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding registry value:",sArgumentValue))
          Call addTest("Registry Value",sArgumentValue)
          bNeedRegObj = True
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If
		    
		  Case "-lu","--local-user","/lu","/local-user"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding local user:",sArgumentValue))
		      Call addTest("Local User",sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If
		    
		  Case "-lg","--local-group","/lg","/local-group"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding local group:",sArgumentValue))
		      Call addTest("Local Group",sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If		    

		  Case "-cq","--custom-query","/cq","/custom-query"
		    If sArgumentValue <> "" Then
		      veryVerbose(formatOutput("Adding custom query:",sArgumentValue))
		      Call addTest("Custom Query",sArgumentValue)
		    Else
		    	fatalError("No value given for: " & sArgument)
		    End If	
		    
		  Case "--result-detail","/result-detail"
		    veryVerbose("Will provide detailed results for custom queries")
		    bResultDetails = True
		    		  
		  '-------------------------  
		  ' Output file options
		  '-------------------------
		  Case "-oX"
		    If sArgumentValue <> "" Then
		    	veryVerbose(formatOutput("XML Output file:",sArgumentValue))
		    	If bOutputFile Then
		    	  fatalError("Can only specify 1 output file.")
		    	Else
		    		sOutputFormat = "XML"
		    		sOutputFile = sArgumentValue
		    		bOutputFile = True
		    	End If
		    Else
		    	'verbose(formatOutput("Ignoring option with no value:",sArgument))
		    	fatalError("No value given for: " & sArgument)
		    End If

      Case "-xsl"
        If sArgumentValue <> "" Then
        	veryVerbose(formatOutput("XSL file:",sArgumentValue))
        	sXSLFile = sArgumentValue
        Else
        	fatalError("No value given for: " & sArgument)
        End If
        
      Case "--web-xsl"
        veryVerbose("Referencing XSL file http://www.winquisitor.org/winquisitor.xsl")
        bWebXSL = True

		  Case "-oC"
		    If sArgumentValue <> "" Then
		    	veryVerbose(formatOutput("CSV Output file:",sArgumentValue))
		    	If bOutputFile Then
		    	  fatalError("Can only specify 1 output file.")
		    	Else
		    		sOutputFormat = "CSV"
		    		sOutputFile = sArgumentValue
		    		bOutputFile = True
		    	End If
		    Else
		    	'verbose(formatOutput("Ignoring option with no value:",sArgument))
		    	fatalError("No value given for: " & sArgument)
		    End If
		    
		  Case "--append-output"
		    iOpenOutputFile = ForAppending
		    bAppend = True
		    
		  '-------------------------  
		  ' Unrecognized argument
		  '-------------------------		    
		  Case Else
			  fatalError(formatOutput("Invalid argument:",sArgument))
		  End Select

		  '----------------------------------------
		  ' End Script arguments
		  '----------------------------------------
		
		Else
			' The first character of the argument is not / or - or the argument length is 1
			fatalError(formatOutput("Invalid argument:",sFullArgument))
		End If
	Next

End Sub	
'*************************************************
' End Sub: processCommandLine
'*************************************************


'*************************************************
'
' Sub: displayUsage
'
' Description: Displays correct usage and options 
'  for the script and then exits.
'*************************************************
Sub displayUsage()
  oStdOut.WriteLine sUsage
  WScript.Quit
End Sub
'*************************************************
' End Sub: displayUsage
'*************************************************

' getUBound code from
'   http://gallery.technet.microsoft.com/ScriptCenter/en-us/ff9a6808-d943-48be-be13-dd53950776ae
'*************************************************
'
' Function: getUBound
'
' Description: Returns the UBound of an array. If
'   the array has no elements returns -1
'*************************************************
Function getUBound(aTempArray)
  Dim iIndex : iIndex = -1
  On Error Resume Next
    iIndex = UBound(aTempArray)
  On Error GoTo 0
  getUBound = iIndex
End Function
'*************************************************
' End Function: getUBound
'*************************************************

'*************************************************
'
' Function: checkService
'
' Description: Checks if the given service is
'   running on the target machine. 
'*************************************************
Function checkService(sTempService)
  Dim cRunningServices
  Dim iItems
  Dim oItem
  Dim sQuery
  
  sQuery = "Select * from Win32_Service Where Name='" & CStr(sTempService) & "' OR DisplayName='" & CStr(sTempService) & "'"
  'debug(sQuery)
  Set cRunningServices = oWMIService.ExecQuery(sQuery) 
  
  iItems = cRunningServices.Count 

  ' If the collection count is greater than zero the service will exist.

  If iItems > 0  Then

    For Each oItem in cRunningServices
      If oItem.State = "Stopped" Then
        checkService = "Installed/Stopped"
        Exit Function
      ElseIf (oItem.State = "Started") Or (oItem.State = "Running") Then
        checkService = "Installed/Running"
        Exit Function
      Else
      	checkService = oItem.State
      	Exit Function
      End If
    Next

  Else
    checkService = "Service Not Installed"
  End If

End Function
'*************************************************
' End Function: checkService
'*************************************************

'*************************************************
'
' Function: checkPatch
'
' Description: Checks if the given patch has
'   been applied to the target machine. 
'*************************************************
Function checkPatch(sTempPatch)
  Dim cPatches
  Dim iItems
  Dim oItem
  Dim sQuery
  
  sQuery = "Select * from win32_quickfixengineering Where HotFixID='" & CStr(sTempPatch) & "'"
  Set cPatches = oWMIService.ExecQuery(sQuery) 
  
  iItems = cPatches.Count 

  ' If the collection count is greater than zero the patch will exist.

  If iItems > 0  Then
    checkPatch = "Installed" 
  Else
    checkPatch = "Not Found"
  End If

End Function
'*************************************************
' End Function: checkPatch
'*************************************************

'*************************************************
'
' Function: checkProcess
'
' Description: Checks if the given process is
'   running. 
'*************************************************
Function checkProcess(sTempProcess)
  Dim cProcesses
  Dim iItems
  Dim oItem
  Dim sQuery
  Dim sPropertyName
  Dim sPropertyValue
  Dim oProperty
  Dim aQueryResults()
  Dim sItemResult
  Dim sOwnerName
  Dim sOwnerDomain
  Dim iReturnCode
  Dim sOwnerSid
  
  sQuery = "Select * from win32_process Where Name='" & CStr(sTempProcess) & "'"
  Set cProcesses = oWMIService.ExecQuery(sQuery) 
  
  iItems = cProcesses.Count 


  If iItems > 0  Then
  	If bResultDetails Then
      For Each oItem in cProcesses
        sPropertyName = ""
        sPropertyValue = ""
        sItemResult = ""
        
        iReturnCode = oItem.GetOwner(sOwnerName,sOwnerDomain)

        If iReturnCode = 0 Then
        	sItemResult = "Owner: " & sOwnerName
        	
        	If IsNull(sOwnerDomain) Then
        		sOwnerDomain = "NULL"
          End If
        		
        	sItemResult = sItemResult & PROPERTY_SEPARATOR & "OwnerDomain: " & sOwnerDomain
        End If

        iReturnCode = oItem.GetOwnerSid(sOwnerSid)

        If iReturnCode = 0 Then
        	sItemResult = sItemResult & PROPERTY_SEPARATOR & "OwnerSID: " & sOwnerSid
        End If

        
        For Each oProperty in oItem.Properties_
          sPropertyName = oProperty.Name
          If IsNull(oProperty.Value) Then
          	sPropertyValue = "NULL"
          ElseIf oProperty.IsArray Then
          	For i = LBound(oProperty.Value) To UBound(oProperty.Value)
          	  sPropertyValue = sPropertyValue & " | " & oProperty.Value
          	Next
          Else
          	sPropertyValue = oProperty.Value
          End If
          'Call addArrayItem(aItemResult,sPropertyName & ": " & sPropertyValue)
          If sItemResult <> "" Then
            sItemResult = sItemResult & PROPERTY_SEPARATOR & sPropertyName & ": " & sPropertyValue
          Else 
          	sItemResult = sPropertyName & ": " & sPropertyValue
          End If
        Next
        Call addArrayItem(aQueryResults,sItemResult)
      Next
      checkProcess = aQueryResults 
      Exit Function
    Else
    	If iItems = 1 Then
    		checkProcess = "1 Instance Running"
    		Exit Function
    	Else
    		checkProcess = iItems & " Instances Running"
    		Exit Function
    	End If
    End If
  Else
    checkProcess = "Not found"
    Exit Function
  End If

End Function
'*************************************************
' End Function: checkProcess
'*************************************************

'*************************************************
'
' Function: checkLocalUser
'
' Description: Checks if the given user account
'   exists on the target machine. 
'*************************************************
Function checkLocalUser(sTempLocalUser)
  Dim cUsers
  Dim iItems
  Dim oItem
  Dim sQuery
  
  sQuery = "Select * from Win32_UserAccount Where Name='" & CStr(sTempLocalUser) & "'"
  Set cUsers = oWMIService.ExecQuery(sQuery) 
  
  iItems = cUsers.Count 

  ' If the collection count is greater than zero the service will exist.

  If iItems > 0  Then
    checkLocalUser = "Exists" 
  Else
    checkLocalUser = "Not Found"
  End If

End Function
'*************************************************
' End Function: checkLocalUser
'*************************************************

'*************************************************
'
' Function: enumLocalGroup
'
' Description: Will list all members of a local
'   group if it exists.
'*************************************************
Function enumLocalGroup(sTempLocalGroup)
  Dim sTempTarget : sTempTarget = sCurrentTarget
  Dim aGroupResults()
  
  If Not bDebug Then On Error Resume Next    		
  
  ' If localhost or 127.0.0.1 is specified as a target, the local group test
  ' will take over 15 seconds. Switch it to .
  If (sTempTarget = "localhost") Or (sTempTarget = "127.0.0.1") Then
    sTempTarget = "."
  End If
      		
  debug("Testing local group " & sTempLocalGroup & " on " & sTempTarget)      		  

  Err.Clear
  If sUsername <> "" Then
  	debug("Connecting to WinNT with alternate credentials")
    Set oGroup = objNS.OpenDSObject("WinNT://" & sTempTarget _
                                      & "/" & sTempLocalGroup & ",group", _
                                      sUsername, sPassword, ADS_SECURE_AUTHENTICATION Or _
                                      ADS_USE_ENCRYPTION)
  Else
   	debug("Connecting to WinNT with current credentials")
    Set oGroup = GetObject("WinNT://" & sTempTarget & "/" & sTempLocalGroup & ",group")
  End If

  If Err <> 0 Then
    enumLocalGroup = "Group not found"
    Exit Function
  Else
    For Each oMember In oGroup.Members
      Call addArrayItem(aGroupResults,oMember.Name)
    Next
  End If
  
  enumLocalGroup = aGroupResults
  
End Function
'*************************************************
' End Function: enumLocalGroup
'*************************************************

'*************************************************
'
' Function: checkFile
'
' Description: Checks if the given file exists
'   on the target machine by querying CIM_Datafile 
'*************************************************
Function checkFile(sTempFile)
  Dim cFiles
  Dim iItems
  Dim oItem
  Dim sQuery
  Dim sFileName
  Dim sFileVersion
  
  debug(formatOutput("Checking file:",sTempFile))
  'If InStr(sFileName,"%") <> 0 Then
  '	oWshShell.ExpandEnvironmentStrings(sFileName)
  'End If
  sFileName = Replace(sTempFile,"\\","\") ' Just in case filename already had double \s
  sFileName = Replace(sFileName,"\","\\")
  sQuery = "Select * from CIM_DataFile where Name = '" & sFileName & "'"

  Set cFiles = oWMIService.ExecQuery(sQuery)
    
  iItems = cFiles.Count 

    ' If the collection count is greater than zero the file exists.
    If iItems > 0  Then
      For Each oItem in cFiles
        ' If the file has version information, include it in the output.
        '   If not, output NULL.
        If IsNull(oItem.Version) Then
        	sFileVersion = "NULL"
        Else
        	sFileVersion = oItem.Version
        End If
        checkFile = "Exists - Version: " & sFileVersion
        Exit Function
      Next
    Else
      checkFile = "Not found"
    End If

End Function
'*************************************************
' End Function: checkFile
'*************************************************


'*************************************************
'
' Function: executeQuery
'
' Description: Executes a user-specified query
'*************************************************
Function executeQuery(sTempQuery)
  Dim cResults
  Dim iItems
  Dim oProperty,oItem
  Dim sQuery
  Dim sPropertyName
  Dim sPropertyValue
  Dim aQueryResults()
  'Dim aItemResult()
  Dim sItemResult
  
  If Not bDebug Then On Error Resume Next
  sQuery = sTempQuery
  Set cResults = oWMIService.ExecQuery(sQuery) 

  iItems = cResults.Count 
  'oStdOut.WriteLine cResults.Count
  If Err <> 0 Then
  	executeQuery = "Invalid query"
  	Exit Function
  End If
  
  ' If the collection count is greater than zero the record was found

  If iItems > 0  Then
  	If bResultDetails Then
      For Each oItem in cResults
        sPropertyName = ""
        sPropertyValue = ""
        sItemResult = ""
        'ReDim aItemResult(0)
        
        For Each oProperty in oItem.Properties_
          sPropertyName = oProperty.Name
          If IsNull(oProperty.Value) Then
          	sPropertyValue = "NULL"
          ElseIf oProperty.IsArray Then
          	For i = LBound(oProperty.Value) To UBound(oProperty.Value)
          	  sPropertyValue = sPropertyValue & " | " & oProperty.Value
          	Next
          Else
          	sPropertyValue = oProperty.Value
          End If
          'Call addArrayItem(aItemResult,sPropertyName & ": " & sPropertyValue)
          If sItemResult <> "" Then
            sItemResult = sItemResult & PROPERTY_SEPARATOR & sPropertyName & ": " & sPropertyValue
          Else 
          	sItemResult = sPropertyName & ": " & sPropertyValue
          End If
        Next
        Call addArrayItem(aQueryResults,sItemResult)
      Next
      executeQuery = aQueryResults 
      Exit Function
    Else
    	If iItems = 1 Then
    		executeQuery = "1 Record"
    		Exit Function
    	Else
    		executeQuery = iItems & " Records"
    		Exit Function
    	End If
    End If
  Else
    executeQuery = "0 Records"
    Exit Function
  End If

End Function
'*************************************************
' End Function: executeQuery
'*************************************************


'*************************************************
'
' Sub: addTarget
'
' Description: Adds another machine to the array
'  of targets
'*************************************************
Sub addTarget(sTarget)
  Dim iTargetArraySize

  iTargetArraySize = getUBound(aTargets) + 1
  ReDim Preserve aTargets(iTargetArraySize)
  aTargets(iTargetArraySize) = Trim(sTarget)
End Sub
'*************************************************
' End Sub: addTarget
'*************************************************


'*************************************************
'
' Sub: addTest
'
' Description: Adds a test to the array, aTests
'*************************************************
Sub addTest(sTempTestType, sTempTestParam)
  
  If iNumTests < 254 Then
    iNumTests = iNumTests + 1
  
    aTests(iNumTests,0) = sTempTestType
    aTests(iNumTests,1) = sTempTestParam
  End If
  
End Sub
'*************************************************
' End Sub: addTest
'*************************************************


'*************************************************
'
' Sub: addArrayItem
'
' Description: Adds an item to an array
'*************************************************
Sub addArrayItem(ByRef aTempArray, sItem)
  Dim iTempArraySize

  iTempArraySize = getUBound(aTempArray) + 1
  ReDim Preserve aTempArray(iTempArraySize)
  aTempArray(iTempArraySize) = Trim(sItem)
End Sub
'*************************************************
' End Sub: addArrayItem
'*************************************************



'*************************************************
'
' Sub: readTargetFile
'
' Description: Reads a file containing a list of 
'  computers to use as targets for the script. The
'  file should contain one computer per line.
'*************************************************
Sub readTargetFile(sTargetFile)
  Dim oTargetFileFSO
  Dim oTargetFile
  Dim sNewTarget
  
  Set oTargetFileFSO = WScript.CreateObject("Scripting.FileSystemObject")
  If Not bDebug Then On Error Resume Next
  If oTargetFileFSO.FileExists(sTargetFile) Then
    Set oTargetFile = oTargetFileFSO.OpenTextFile(sTargetFile,ForReading)
    If Err <> 0 Then
    	fatalError("Could not open file: " & sTargetFile & " - " & Err.Description)
    Else
      Do until oTargetFile.AtEndOfStream
      	sNewTarget = oTargetFile.Readline
      	Call addArrayItem(aTargets,sNewTarget)
      	debug(formatOutput("Adding target from file",sNewTarget))
      Loop
      oTargetFile.Close
    End If
   Else
   	fatalError(sTargetFile & " does not exist")
   End If
  On Error GoTo 0
End Sub
'*************************************************
' End Sub: readTargetFile
'*************************************************


Sub writeHeaders(ByRef aTempArray, bNewLine)
  Dim sHeader
  
  For Each sHeader in aTempArray 
    oOutputFile.Write "," & sHeader
    Call increaseTestCount()
  Next
  
  If bNewLine Then
  	oOutputFile.Write vbCRLF
  End If
End Sub

Sub increaseTestCount()
  
  iTestCount = iTestCount + 1

End Sub

'*************************************************
'
' Sub: removeArrayDuplicates
'
' Description: Removes duplicate items from an
'   array.
'*************************************************
Sub removeArrayDuplicates(ByRef aTempArray)
  Dim oDictionary
  Dim sItem
  Dim i : i = 0
  
  If getUBound(aTempArray) > -1 Then
    Set oDictionary = CreateObject("Scripting.Dictionary")
    oDictionary.RemoveAll
    oDictionary.CompareMode = 0

    For Each sItem In aTempArray
      If Len(Trim(sItem)) > 0 Then
        If Not oDictionary.Exists(Trim(sItem)) Then
          oDictionary.Add Trim(sItem), Trim(sItem)
        Else
        	debug(formatOutput("Removing array item:",sItem))
        End If
      End If
    Next
  
    Redim aTempArray(UBound(oDictionary.Items))
    For Each sItem in oDictionary.Items
      aTempArray(i) = sItem
      i = i + 1
    Next
    'aTempArray = oDictionary.Items
    Set oDictionary = Nothing
    
  End If
  
End Sub
'*************************************************
' End Sub: removeArrayDuplicates
'*************************************************


Function checkRegistryKey(sTempRegistryKey)
  Dim sRootKey
  Dim sKeyPath
  Dim sEntryName
  Dim hRootKey
  Dim iFirstSlash
  Dim iLastSlash
  Dim aSubKeys
  
  If Not IsObject(oReg) Then
  	checkRegistryKey = "Could not connect to registry"
    Exit Function
  End If
  
  iFirstSlash = InStr(sTempRegistryKey,"\")
  'debug(formatOutput("First slash",iFirstSlash))
  iLastSlash = InStrRev(sTempRegistryKey,"\")
  'debug(formatOutput("Last slash",iLastSlash))
  If (iFirstSlash < 1) Then
  	checkRegistryKey = "Invalid registry key"
  	Exit Function
  End If
  
  sRootKey = Left(sTempRegistryKey,iFirstSlash-1)
  'sEntryName = Mid(sTempRegistryItem,iLastSlash+1)
  'sKeyPath = Mid(sTempRegistryItem,iFirstSlash+1,iLastSlash-iFirstSlash)
  sKeyPath = Mid(sTempRegistryKey,iFirstSlash+1)
  debug(formatOutput("Root key",sRootKey))
  debug(formatOutput("Key path",sKeyPath))
  'debug(formatOutput("Entry name",sEntryName))
  
  Select Case sRootKey
  	Case "HKLM","HKEY_LOCAL_MACHINE"
  	  hRootKey = HKEY_LOCAL_MACHINE
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKCU","HKEY_CURRENT_USER"
  	  hRootKey = HKEY_CURRENT_USER
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKCR","HKEY_CLASSES_ROOT"
  	  hRootKey = HKEY_CLASSES_ROOT
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKU","HKEY_USERS"
  	  hRootKey = HKEY_USERS
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKCC","HKEY_CURRENT_CONFIG"
  	  hRootKey = HKEY_CURRENT_CONFIG
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKDD","HKEY_DYN_DATA"
  	  hRootKey = HKEY_DYN_DATA
  	  debug(formatOutput("Root key",hRootKey))
  	Case Else
  	  checkRegistryKey = "Invalid registry key - unrecognized root key: " & sRootKey
  	  Exit Function
  End Select
  
  If oReg.EnumKey(hRootKey, sKeyPath, aSubKeys) = 0 Then
    checkRegistryKey = "Exists"
    Exit Function
  End If
  
  checkRegistryKey = "Not found"
  
End Function



Function checkRegistryValue(sTempRegistryValue)
  Dim sRootKey
  Dim sKeyPath
  Dim sEntryName
  Dim sSubKey
  Dim sSubKeyPath
  Dim sValueName
  Dim sValue
  Dim hRootKey
  Dim iFirstSlash
  Dim iLastSlash
  Dim i
  Dim aSubKeys
  Dim aValueNames
  Dim aValues
  Dim aTypes
  
  If Not IsObject(oReg) Then
  	checkRegistryValue = "Could not connect to registry"
    Exit Function
  End If  
  
  iFirstSlash = InStr(sTempRegistryValue,"\")
  debug(formatOutput("First slash",iFirstSlash))
  iLastSlash = InStrRev(sTempRegistryValue,"\")
  debug(formatOutput("Last slash",iLastSlash))
  If (iFirstSlash < 1) Then
  	checkRegistryValue = "Invalid registry key"
  	Exit Function
  End If
  
  sRootKey = Left(sTempRegistryValue,iFirstSlash-1)
  sEntryName = Mid(sTempRegistryValue,iLastSlash+1)
  sKeyPath = Mid(sTempRegistryValue,iFirstSlash+1,iLastSlash-iFirstSlash-1)

  debug(formatOutput("Root key",sRootKey))
  debug(formatOutput("Key path",sKeyPath))
  debug(formatOutput("Entry name",sEntryName))
  
  Select Case sRootKey
  	Case "HKLM","HKEY_LOCAL_MACHINE"
  	  hRootKey = HKEY_LOCAL_MACHINE
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKCU","HKEY_CURRENT_USER"
  	  hRootKey = HKEY_CURRENT_USER
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKCR","HKEY_CLASSES_ROOT"
  	  hRootKey = HKEY_CLASSES_ROOT
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKU","HKEY_USERS"
  	  hRootKey = HKEY_USERS
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKCC","HKEY_CURRENT_CONFIG"
  	  hRootKey = HKEY_CURRENT_CONFIG
  	  debug(formatOutput("Root key",hRootKey))
  	Case "HKDD","HKEY_DYN_DATA"
  	  hRootKey = HKEY_DYN_DATA
  	  debug(formatOutput("Root key",hRootKey))
  	Case Else
  	  checkRegistryKey = "Invalid registry key - unrecognized root key: " & sRootKey
  	  Exit Function
  End Select
  
        sSubKeyPath = sKeyPath 
        oReg.EnumValues hRootKey, sSubKeyPath, aValueNames, aTypes

        If getUBound(aValueNames) > -1 Then
        	For i = LBound(aValueNames) To UBound(aValueNames)
        	  sValueName = aValueNames(i)
        	  If sValueName = sEntryName Then
        	  Select Case aTypes(i)
        	  	Case REG_SZ      
                oReg.GetStringValue hRootKey, sSubKeyPath, sValueName, sValue
                checkRegistryValue = "(REG_SZ) = " & sValue
                Exit Function
              Case REG_EXPAND_SZ
                oReg.GetExpandedStringValue hRootKey, sSubKeyPath, sValueName, sValue
                checkRegistryValue = "(REG_EXPAND_SZ) = " & sValue
                Exit Function
              Case REG_BINARY
                Dim sBytes : sBytes = ""
                Dim aBytes
                Dim uByte
                
                oReg.GetBinaryValue hRootKey, sSubKeyPath, sValueName, aBytes
                For Each uByte in aBytes
                  sBytes = sBytes & Hex(uByte) & " "
                Next
                checkRegistryValue = "(REG_BINARY) = " & sBytes
                Exit Function
              Case REG_DWORD
                Dim uValue
                
                oReg.GetDWORDValue hRootKey, sSubKeyPath, sValueName, uValue
                checkRegistryValue = "(REG_DWORD) = " & CStr(uValue)				  
                Exit Function
              Case REG_MULTI_SZ
                Dim sReturnValue
                
                oReg.GetMultiStringValue hRootKey, sSubKeyPath, sValueName, aValues				  				
                sReturnValue = "(REG_MULTI_SZ) ="
                For Each sValue in aValues
                  sReturnValue = sReturnValue & sValue 
                Next
                checkRegistryValue = sReturnValue
                Exit Function
        	  End Select
        	  End If
          Next
        Else
        	checkRegistryValue = "Subkey Not Found"
        	Exit Function
        End If

  
  checkRegistryValue = "Registry Key Not found"
  
End Function



' PingStatus based mostly on the following two sites
'   http://www.tek-tips.com/faqs.cfm?fid=4871
'   http://www.microsoft.com/technet/scriptcenter/resources/scriptshop/shop1205.mspx
Function PingStatus(sComputer)

    If Not bDebug Then On Error Resume Next
    Dim oLocalWMIService, cPings, oPing
    
    Set oLocalWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set cPings = oLocalWMIService.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & sComputer & "'")
    
    For Each oPing in cPings
        Select Case oPing.StatusCode
            Case 0 PingStatus = "Success"
            Case 11001 PingStatus = "Status code 11001 - Buffer Too Small"
            Case 11002 PingStatus = "Status code 11002 - Destination Net Unreachable"
            Case 11003 PingStatus = "Status code 11003 - Destination Host Unreachable"
            Case 11004 PingStatus = "Status code 11004 - Destination Protocol Unreachable"
            Case 11005 PingStatus = "Status code 11005 - Destination Port Unreachable"
            Case 11006 PingStatus = "Status code 11006 - No Resources"
            Case 11007 PingStatus = "Status code 11007 - Bad Option"
            Case 11008 PingStatus = "Status code 11008 - Hardware Error"
            Case 11009 PingStatus = "Status code 11009 - Packet Too Big"
            Case 11010 PingStatus = "Status code 11010 - Request Timed Out"
            Case 11011 PingStatus = "Status code 11011 - Bad Request"
            Case 11012 PingStatus = "Status code 11012 - Bad Route"
            Case 11013 PingStatus = "Status code 11013 - TimeToLive Expired Transit"
            Case 11014 PingStatus = "Status code 11014 - TimeToLive Expired Reassembly"
            Case 11015 PingStatus = "Status code 11015 - Parameter Problem"
            Case 11016 PingStatus = "Status code 11016 - Source Quench"
            Case 11017 PingStatus = "Status code 11017 - Option Too Big"
            Case 11018 PingStatus = "Status code 11018 - Bad Destination"
            Case 11032 PingStatus = "Status code 11032 - Negotiating IPSEC"
            Case 11050 PingStatus = "Status code 11050 - General Failure"
            Case Else PingStatus = "Status code " & oPing.StatusCode & " - Unable to determine cause of failure."
        End Select
    Next

End Function





Function formatOutput(sString,sValue)
  Dim iStringPadding
  Dim sPaddingCharacter
  
  iStringPadding = 40 - Len(sString)
  sPaddingCharacter = "."
  
  If iStringPadding > -1 Then
    formatOutput = CStr(sString) & String(iStringPadding,sPaddingCharacter) & CStr(sValue)
  Else
  	formatOutput = CStr(sString) & vbCRLF & Space(5) & String(35,sPaddingCharacter) & CStr(sValue)
  End If

End Function



Sub debug(sText)
	If bDebug Then
		oStdOut.WriteLine "DEBUG: (" & Round(Timer - iTimer, 4) & "s) " & CStr(sText)
	End If
End Sub

Sub output(sText)
	If Not bQuiet Then
		oStdOut.WriteLine CStr(sText)
	End If
End Sub

Sub verbose(sText)
	If bVerbose Then
		oStdOut.WriteLine CStr(sText)
	End If
End Sub

Sub veryVerbose(sText)
	If bVeryVerbose Then
		oStdOut.WriteLine CStr(sText)
	End If
End Sub

Sub displayError(sText)
  If Not bQuiet Then
  	oStdErr.WriteLine CStr(sText)
  End If
End Sub

Sub fatalError(sText)
  oStdErr.Write vbCRLF & "ERROR: " & CStr(sText) & vbCRLF
  WScript.Quit
End Sub

Sub recordResult(sTempTarget,sTestType,sTest,sTestResult)
  Dim sResultItem : sResultItem = ""
  Dim bFirstItem  : bFirstItem = True
  
  If IsArray(sTestResult) Then
    output(Space(3) & sTest & ":")
    If getUBound(sTestResult) > -1 Then
    	For Each sResultItem in sTestResult
    	  output(formatOutput(Space(6),sResultItem))
    	Next
    End If
  Else
    output(formatOutput(Space(3) & sTest & ":",sTestResult))
  End If
  
  If  bOutputFile Then
  	Select Case sOutputFormat
  		Case "XML"
  		  Call openXMLElement("test")
  		  Call fullXMLElement("type",sTestType)
  		  Call fullXMLElement("value",sTest)
  		  If IsArray(sTestResult) Then
  		  	For Each sResultItem In sTestResult
  		  	  Call fullXMLElement("result",sResultItem)
  		  	Next
  		  Else
  		  	Call fullXMLElement("result",sTestResult)
  		  End If
  		  Call closeXMLElement("test")
  		Case "CSV"
  		  ' Output format should be: TestType,TestParam,Result
  		  ' If sTestResult is an array need to combine the elements and use
  		  '   RECORD_SEPARATOR to separate each one
  		  If IsArray(sTestResult) Then
  		  	oOutputFile.Write "," & sTestType & "," & sTest & ","
  		  	For Each sResultItem In sTestResult
  		  	  If bFirstItem Then
  		  	    oOutputFile.Write sResultItem
  		  	    bFirstItem = False
  		  	  Else
  		  	  	oOutputFile.Write RECORD_SEPARATOR & sResultItem
  		  	  End If
  		  	Next
  		  Else
  		    oOutputFile.Write "," & sTestType & "," & sTest & "," & sTestResult
  		  End If
  	End Select
  End If
  
End Sub

Sub recordNewTarget(sTempTarget)
  If bOutputFile Then
    Select Case sOutputFormat
    	Case "XML"
        Call openXMLElement("target")
    	  Call fullXMLElement("computer",sCurrentTarget)
    	Case "CSV"
    	  oOutputFile.Write sTempTarget
    End Select
  End If
End Sub

Sub endCurrentTarget(sTempTarget)
  If bOutputFile Then
    Select Case sOutputFormat
    	Case "XML"
    	  Call closeXMLElement("target")
    	Case "CSV"
    	  oOutputFile.Write vbNewLine
    End Select
  End If
End Sub

Sub prepareOutputFile()
  If bOutputFile Then
    Select Case sOutputFormat
      Case "XML"
        If Not bAppend Then
    	    Call beginXML
    	    Call openXMLElement("winquisitor_audit")
    	  End If
    	  Call openXMLElement("scan")
    	  sScanInfo = WScript.ScriptName
    	  For i = 0 To WScript.Arguments.Count - 1
    	    sScanInfo = sScanInfo & " " & WScript.Arguments(i)
    	  Next
    	  Call fullXMLElement("scan_info",sScanInfo)
        Call fullXMLElement("start_date",Date)
    	  Call fullXMLElement("start_time",Time)  
      Case "CSV"
        If Not bAppend Then
          oOutputFile.Write "Computer" & "," & "Connection"
          For i = 0 To iNumTests
 	          oOutputFile.Write ",TestType,Parameter,Result"
 	        Next
 	        oOutputFile.Write vbCRLF
        End If
    End Select
  End If
End Sub

Sub recordConnectFailure(sMessage)
  If bOutputFile Then 
    Select Case sOutputFormat
      Case "XML"
        Call fullXMLElement("connection","Failed")
        Call fullXMLElement("error",sMessage)  
      Case "CSV"
        oOutputFile.Write "," & sMessage & "," & String(iTestCount - 1,",")
    End Select
  End If
End Sub

Sub recordConnectSuccess()
  If bOutputFile Then
    Select Case sOutputFormat
      Case "XML"
        Call fullXMLElement("connection","Success")  
      Case "CSV"
        oOutputFile.Write "," & "Success"
    End Select
  End If
End Sub

Sub fullXMLElement(sElement,sXMLValue)
  oOutputFile.Write(vbNewLine & "<" & sElement & ">")
  oOutputFile.Write(sXMLValue)
  oOutputFile.Write("</" & sElement & ">")
End Sub

Sub openXMLElement(sElement)
  oOutputFile.Write(vbNewLine & "<" & sElement & ">")
End Sub

Sub closeXMLElement(sElement)
  oOutputFile.Write(vbNewLine & "</" & sElement & ">")
End Sub

Sub writeXMLValue(sXMLValue)
  oOutputFile.Write(sElement)
End Sub

Sub beginXML()
  oOutputFile.WriteLine("<?xml version='1.0' encoding='ISO-8859-1'?>")
  If sXSLFile <> "" Then
  	oOutputFile.Write("<?xml-stylesheet type='text/xsl' href='" & sXSLFile & "'?>")
  	Exit Sub
  End If 
  	
  If bWebXSL Then
  	oOutputFile.Write("<?xml-stylesheet type='text/xsl' href='http://www.winquisitor.org/winquisitor.xsl'?>")
  	Exit Sub
  End If

  If (Not bWebXSL) And (sXSLFile = "") Then  
    oOutputFile.Write("<?xml-stylesheet type='text/xsl' href='winquisitor.xsl'?>")
    Exit Sub
  End If
  
End Sub
