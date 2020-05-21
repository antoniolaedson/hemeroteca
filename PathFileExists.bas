Attribute VB_Name = "modPathFileExists"
Option Explicit
Public Declare Function Existe Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
