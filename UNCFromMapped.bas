'------------------------------------------------------------------------------
'Purpose  : Converts a (local) drive mapping to a network share/drive to its UNC path notation
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 2013-2017
'   Source: -
'  Changed:
'------------------------------------------------------------------------------
#Compile Exe ".\UNCFromMappedDrive.exe"
#Dim All

#Debug Error Off
#Tools Off

DefLng A-Z

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 1

' Version Resource information
#Include ".\UNCFromMappedRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include Once "win32api.inc"
#Include "sautilcc.inc"
'------------------------------------------------------------------------------
'*** Variabels ***
'------------------------------------------------------------------------------
'==============================================================================

Function PBMain () As Long
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 14.02.2017
'           - Code refactoring for Github publication
'------------------------------------------------------------------------------
   Local sTemp As String, dwRet As Dword

   ' Application intro
   ConHeadline "UNCFromMappedDrive", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2013-2017", $COMPANY_NAME
   Print ""

   If Len(Trim$(Command$)) < 1 Or InStr(Command$, "/?") > 0 Then
      ShowHelp
      Exit Function
   End If

   sTemp = Command$
   sTemp = Remove$(sTemp, $Dq)

   StdOut UNCPathFromDriveLetter(sTemp, dwRet)

   If dwRet <> %NO_ERROR Then
      Function = 255
   End If

End Function
'------------------------------------------------------------------------------

Function UNCPathFromDriveLetter(ByVal sPath As String, ByRef dwError As Dword, _
   Optional ByVal lDriveOnly As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a fully qulified UNC path location from a (mapped network)
'           drive letter/share
'
'Prereq.  : -
'Parameter: sPath       - Path to resolve
'           dwError     - ByRef(!), Returns the error code from the Win32 API, if any
'           lDriveOnly  - If True, return only the drive letter
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 17.07.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   ' 32-bit declarations:
   Local sTemp As String
   Local szDrive As AsciiZ * 3, szRemoteName As AsciiZ * 1024
   Local lSize, lStatus As Long

   ' The size used for the string buffer. Adjust this if you
   ' need a larger buffer.
   Local lBUFFER_SIZE As Long
   lBUFFER_SIZE = 1024

   If Len(sPath) > 2 Then
      sTemp = Mid$(sPath, 3)
      szDrive = Left$(sPath, 2)
   Else
      szDrive = sPath
   End If

   ' Return the UNC path (\\Server\Share).
   lStatus = WNetGetConnectionA(szDrive, szRemoteName, lBUFFER_SIZE)

   ' Verify that the WNetGetConnection() succeeded. WNetGetConnection()
   ' returns 0 (NO_ERROR) if it successfully retrieves the UNC path.
   If lStatus = %NO_ERROR Then

      If IsTrue(lDriveOnly) Then

         ' Display the UNC path.
         UNCPathFromDriveLetter = Trim$(szRemoteName, Any $Nul & $WhiteSpace)

      Else

         UNCPathFromDriveLetter = Trim$(szRemoteName, Any $Nul & $WhiteSpace) & sTemp

      End If

   Else

      ' Return the original filename/path unaltered
      UNCPathFromDriveLetter = sPath

   End If

   dwError = lStatus

End Function
'==============================================================================

Sub ShowHelp
'------------------------------------------------------------------------------
'Purpose  : Help screen
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   StdOut "UNCFromMappedDrive"
   StdOut "------------------"
   StdOut "UNCFromMappedDrive converts a (local) drive mapping to a network share/drive to its UNC path notation."
   StdOut "i.e. W:\MyData will be resolved to \\MyServer\MyNetworkShare\MyNetWorkData\MyData."
   StdOut ""
   StdOut "Usage:   UNCFromMappedDrive <mapped drive/folder>.
   StdOut "i.e.     UNCFromMappedDrive W: (drive only)"
   StdOut "         - or -"
   StdOut "         UNCFromMappedDrive W:\MyData (drive & path)"
   StdOut ""

End Sub
'---------------------------------------------------------------------------
