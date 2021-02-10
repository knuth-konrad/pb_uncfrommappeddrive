'------------------------------------------------------------------------------
'Purpose  : Converts a (local) drive mapping to a network share/drive to its UNC path notation
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 2013-2017
'   Source: -
'  Changed: 15.05.2017
'           - #Break On to prevent console window property's menu issue
'           - Application manifest added
'------------------------------------------------------------------------------
#Compile Exe ".\UNCFromMappedDrive.exe"
#Dim All

#Break On
#Debug Error Off
#Tools Off

DefLng A-Z

%VERSION_MAJOR = 1
%VERSION_MINOR = 0
%VERSION_REVISION = 3

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

Macro TextToClip(st)

   MacroTemp hMem, pMem
   Dim hMem    As Dword  'handle to globally allocated memory
   Dim pMem    As Dword  'pointer to globally allocated memory

   If OpenClipboard(0) Then
      EmptyClipboard

      hMem = GlobalAlloc( %GMEM_MOVEABLE Or %GMEM_DDESHARE, Len( st ) + 1 )

      If hMem Then
         pMem = GlobalLock( hMem )
         If pMem Then
            Poke$ pMem, st & $Nul
            GlobalUnlock hMem
            SetClipboardData %CF_Text, hMem
         End If
      End If 'pMem

   CloseClipboard

   End If 'OpenClipBoard

End Macro 'TextToClip
'------------------------------------------------------------------------------

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
'           04.06.2019
'           - Copy the result to the clipboard for easier usage
'------------------------------------------------------------------------------
   Local sTemp As String, sResult As String
   Local dwRet As Dword

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

   sResult = UNCPathFromDrive(sTemp, dwRet)

   Con.StdOut sResult
   TextToClip(sResult)

   If dwRet <> %NO_ERROR Then
      Function = 255
   End If

End Function
'------------------------------------------------------------------------------

Function UNCPathFromDrive(ByVal sPath As String, ByRef dwError As Dword, _
   Optional ByVal lDriveOnly As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a fully qualified UNC path location from a (mapped network)
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
         UNCPathFromDrive = Trim$(szRemoteName, Any $Nul & $WhiteSpace)

      Else

         UNCPathFromDrive = Trim$(szRemoteName, Any $Nul & $WhiteSpace) & sTemp

      End If

   Else

      ' Return the original filename/path unaltered
      UNCPathFromDrive = sPath

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

   Con.StdOut "UNCFromMappedDrive"
   Con.StdOut "------------------"
   Con.StdOut "UNCFromMappedDrive converts a (local) drive mapping to a network share/drive to its UNC path notation."
   Con.StdOut "i.e. W:\MyData will be resolved to \\MyServer\MyNetworkShare\MyNetWorkData\MyData."
   Con.StdOut ""
   Con.StdOut "Usage:   UNCFromMappedDrive <mapped drive/folder>.
   Con.StdOut "i.e.     UNCFromMappedDrive W: (drive only)"
   Con.StdOut "         - or -"
   Con.StdOut "         UNCFromMappedDrive W:\MyData (drive & path)"
   Con.StdOut ""

End Sub
'---------------------------------------------------------------------------
