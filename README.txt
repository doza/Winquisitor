winquisitor.vbs v0.1.5 ( http://winquisitor.org ) 

AUTHOR: Mike Cardosa
        http://twitter.com/doza
LAST UPDATED: September 22, 2010


DESCRIPTION:
=====================

Winquisitor aims to simplify the tasks that Windows administrators must perform
by providing a simple way to gather information from a number of Windows 
systems, reducing custom script development.


DISCLAIMER:
=====================

The author makes no representations about the suitability
of this software for any purpose.  This software is provided
AS IS and without any express or implied warranties, 
including, without limitation, the implied warranties of 
merchantability and fitness for a particular purpose. The
entire risk arising out of the use or performance of this script 
and documentation remains with you. In no event shall the author,
or anyone else involved in the creation, production, or delivery
of the scripts be liable for any damages whatsoever (including,
without limitation, damages for loss of business profits, 
business interruption, loss of business information, or other
pecuniary loss) arising out of the use of or inability to use
the script or documentation, even if the author has been
advised of the possibility of such damages.


INSTALLATION:
=====================

Simply extract winquisitor.vbs to any local directory. 

If you wish to view XML in a browser formatted using the included 
winquisitor.xsl, copy winquisitor.xsl to the report directory or specify 
the path to the XSL file on the command line with the -xsl option.


USAGE:
=====================

cscript [ //nologo ] winquisitor.vbs [ -h|--help ]

cscript [ //nologo ] winquisitor.vbs { test(s) } [ output ] { target specification }


PARAMETERS:
=====================

 OUTPUT:
 --------------------
  -h,--help                    Display this usage screen
  -v                           Enable verbose output
  -vv                          Enable very verbose output
  -d,--debug                   Enable debugging output
  -q,--quiet                   Suppress output
  -oC:file                     Output CSV results to the given file
  -oX:file                     Output XML results to the given file
  -xsl:file                    Reference the given XSL document in the
                                 XML output file instead of the default
                                 winquisitor.xsl
  --web-xsl                    Reference the XSL file hosted on winquisitor.org
                                 in the XML output file instead of the
                                 default winquisitor.xsl
                                 Note: This will not work in Firefox because
                                 FF will not parse XSL files from a different
                                 scope than the XML file.
  --append-output              Append to the given output file instead of 
                                 overwriting

 TARGET SPECIFICATION:
 --------------------
  -t,--target:computer         Add the given computer to the list of computers
                                 to test
  -T,--target-file:file        Read targets from the given file
                                 (one target per line)
  -np,--no-ping                Do not ping targets before trying to connect
  -u,--username:username       Connect to targets with the given username
  -p,--password:password       Connect to targets with the given password
                                 If a username was given and a password was
                                 not specified, then the user will be prompted
                                 for a password.

 TESTS:
 --------------------
  -f,--file:file               Test the existence and version of the given file
  -s,--service:service         Test the state of the given service
  -pa,--patch:patch            Test whether a given patch has been applied
  -pr,--process:process        Test whether or not a process is running
  -rk,--registry-key:key       Test the existence and/or value of the
                                 given registry key
  -rv,--regisry-value:value    Test the given registry value
  -lu,--local-user:username    Test the existence of the given user
  -lg,--local-group:groupname  Enumerate the members of the given local group
  -cq,--custom-query:query     WMI query against the CIMV2 namespace
  --result-detail              Provide detailed results instead of a summary.
                                 Any properties and values will be enumerated.


EXAMPLES:
=====================

 EXAMPLE 1:
 --------------------
  Test for the Alerter service on machines 192.168.1.10 and 192.168.1.11
    and record results in XML format to results.xml

   winquisitor.vbs -t:192.168.1.10 -t:192.168.1.11 -s:Alerter -oX:results.xml


 EXAMPLE 2:
 --------------------
  Test for the existence of the file "C:\Windows\system32\evil.exe" and
    the running process trojan.exe against 192.168.1.10, 192.168.1.1, and all
    hosts listed in targets.txt. Record detailed results in XML format
    to results.xml

   winquisitor.vbs -t:192.168.1.10 -t:192.168.1.11 -T:targets.txt
     -f:"C:\Windows\system32\evil.exe" -p:"trojan.exe" -oX:results.xml
     --result-detail


 EXAMPLE 3:
 --------------------
  Check for patch KB890046 and run a custom query against 192.168.1.11
    displaying detailed results. Do not ping the target first. Append the
    results in CSV format to results.csv

   winquisitor.vbs -t:192.168.1.11 -np -pa:KB890046 -oC:results.csv
     -cq:"select caption from win32_useraccount" --result-detail --append-output
     