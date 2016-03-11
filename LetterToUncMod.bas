Attribute VB_Name = "LetterToUncMod"
 Option Explicit

      Private Const RESOURCETYPE_ANY = &H0
      Private Const RESOURCE_CONNECTED = &H1

      Private Type NETRESOURCE
         dwScope As Long
         dwType As Long
         dwDisplayType As Long
         dwUsage As Long
         lpLocalName As Long
         lpRemoteName As Long
         lpComment As Long
         lpProvider As Long
      End Type

      Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
         "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, _
         ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long

      Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" _
      (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long

      Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long

      Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
         (ByVal lpString As Any) As Long

      Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
         (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long



      Function LetterToUNC(DriveLetter As String) As String
         Dim hEnum As Long
         Dim NetInfo(1023) As NETRESOURCE
         Dim entries As Long
         Dim nStatus As Long
         Dim LocalName As String
         Dim UNCName As String
         Dim i As Long
         Dim r As Long

         ' Begin the enumeration
         nStatus = WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0&, ByVal 0&, hEnum)

        'if not found return original drive letter
         LetterToUNC = DriveLetter

         'Check for success from open enum
         If ((nStatus = 0) And (hEnum <> 0)) Then
            ' Set number of entries
            entries = 1024

            ' Enumerate the resource
            nStatus = WNetEnumResource(hEnum, entries, NetInfo(0), CLng(Len(NetInfo(0))) * 1024)

            ' Check for success
            If nStatus = 0 Then
               For i = 0 To entries - 1
                  ' Get the local name
                  LocalName = ""
                  If NetInfo(i).lpLocalName <> 0 Then
                     LocalName = Space(lstrlen(NetInfo(i).lpLocalName) + 1)
                     r = lstrcpy(LocalName, NetInfo(i).lpLocalName)
                  End If

                  ' Strip null character from end
                  If Len(LocalName) <> 0 Then
                     LocalName = Left(LocalName, (Len(LocalName) - 1))
                  End If

                  If UCase$(LocalName) = UCase$(DriveLetter) Then
                     ' Get the remote name
                     UNCName = ""
                     If NetInfo(i).lpRemoteName <> 0 Then
                        UNCName = Space(lstrlen(NetInfo(i).lpRemoteName) + 1)
                        r = lstrcpy(UNCName, NetInfo(i).lpRemoteName)
                     End If

                     ' Strip null character from end
                     If Len(UNCName) <> 0 Then
                        UNCName = Left(UNCName, (Len(UNCName) - 1))
                     End If

                     ' Return the UNC path to drive
                     LetterToUNC = UNCName

                     ' Exit the loop
                     Exit For
                  End If
               Next i
            End If
         End If

         ' End enumeration
         nStatus = WNetCloseEnum(hEnum)
      End Function




