Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" _
 Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function fcnUTILITIES_GetPrivateProfileString(sSection As String, sItem As String, sFileName As String)
  Dim nReturnLen%
  Dim sReturnString As String * 60
  nReturnLen = GetPrivateProfileString(sSection, sItem, "", sReturnString, Len(sReturnString), sFileName)
  fcnUTILITIES_GetPrivateProfileString = Left$(sReturnString, nReturnLen)
End Function

