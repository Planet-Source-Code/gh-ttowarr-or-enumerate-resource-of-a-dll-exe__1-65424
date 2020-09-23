Attribute VB_Name = "ModRes"
Option Explicit

'Public APIs
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Any, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As Any, ByVal lpName As Any, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

'Private variables
Private ResMod As Long
Private UserTreeView As TreeView

' Public Predefined Resource Types
Public Enum ResTypes
   RT_CURSOR = 1&
   RT_BITMAP = 2&
   RT_ICON = 3&
   RT_MENU = 4&
   RT_DIALOG = 5&
   RT_STRING = 6&
   RT_FONTDIR = 7&
   RT_FONT = 8&
   RT_ACCELERATOR = 9&
   RT_RCDATA = 10&
   RT_MESSAGETABLE = 11&
   RT_GROUP_CURSOR = 12&
   RT_GROUP_ICON = 14&
   RT_VERSION = 16&
   RT_DLGINCLUDE = 17&
   RT_PLUGPLAY = 19&
   RT_VXD = 20&
   RT_ANICURSOR = 21&
   RT_ANIICON = 22&
   RT_HTML = 23&
End Enum

'Public callback type
Public Function CallBack_EnumResourceTypes(ByVal hModule As Long, ByVal lpType As Long, ByVal lParam As Long) As Long
EnumResourceNames hModule, lpType, AddressOf CallBack_EnumResourceNames, lParam
CallBack_EnumResourceTypes = True
End Function

'Public callback name
Public Function CallBack_EnumResourceNames(ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal lParam As Long) As Long
EnumResourceLanguages hModule, lpType, lpName, AddressOf CallBack_EnumResourceLanguages, lParam
CallBack_EnumResourceNames = True
End Function

'Public callback language
Public Function CallBack_EnumResourceLanguages(ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLang As Long, ByVal lParam As Long) As Long
Dim ResTypeName As String
Dim ResNameName As String
Dim i As Integer

ResTypeName = DecodeResTypeName(lpType)
ResNameName = DecodeResNameName(lpName)
With FrmMain.tView
        For i = 1 To .Nodes.Count
            If .Nodes(i).Key = "M" & ResTypeName Then 'Lptype is already
                .Nodes.Add "M" & ResTypeName, tvwChild, "I" & ResTypeName & ", " & ResNameName, ResNameName, 2
                Exit Function
            End If
        Next i
        'Lptype is not already
        .Nodes.Add , tvwChild, "M" & ResTypeName, ResTypeName, 1
        .Nodes.Add "M" & ResTypeName, tvwChild, "I" & ResTypeName & ", " & ResNameName, ResNameName, 2
End With

CallBack_EnumResourceLanguages = True
End Function

'Private Select Resname standard
Private Function ResTypeName(ByVal ResType As ResTypes) As String
   Select Case ResType
      Case RT_ACCELERATOR
         ResTypeName = "Accelerator table"
      Case RT_ANICURSOR
         ResTypeName = "Animated cursor"
      Case RT_ANIICON
         ResTypeName = "Animated icon"
      Case RT_BITMAP
         ResTypeName = "Bitmap resource"
      Case RT_CURSOR
         ResTypeName = "Hardware-dependent cursor resource"
      Case RT_DIALOG
         ResTypeName = "Dialog box"
      Case RT_DLGINCLUDE
         ResTypeName = "Header file that contains menu and dialog box #define statements"
      Case RT_FONT
         ResTypeName = "Font resource"
      Case RT_FONTDIR
         ResTypeName = "Font directory resource"
      Case RT_GROUP_CURSOR
         ResTypeName = "Hardware-independent cursor resource"
      Case RT_GROUP_ICON
         ResTypeName = "Hardware-independent icon resource"
      Case RT_HTML
         ResTypeName = "HTML document"
      Case RT_ICON
         ResTypeName = "Hardware-dependent icon resource"
      Case RT_MENU
         ResTypeName = "Menu resource"
      Case RT_MESSAGETABLE
         ResTypeName = "Message-table entry"
      Case RT_PLUGPLAY
         ResTypeName = "Plug and play resource"
      Case RT_RCDATA
         ResTypeName = "Application-defined resource (raw data)"
      Case RT_STRING
         ResTypeName = "String-table entry"
      Case RT_VERSION
         ResTypeName = "Version resource"
      Case RT_VXD
         ResTypeName = "VXD"
      Case Else
         ResTypeName = "User-defined custom resource"
   End Select
End Function

'Private Decode Resourcetype name
Private Function DecodeResTypeName(ByVal lpszValue As Long) As String
Dim Buffer() As Byte
Dim nLen As Long

If HiWord(lpszValue) Then
    nLen = lstrlenA(ByVal lpszValue)
    ReDim Buffer(0 To (nLen - 1)) As Byte
    CopyMemory Buffer(0), ByVal lpszValue, nLen
    DecodeResTypeName = StrConv(Buffer, vbUnicode)
Else
    DecodeResTypeName = ResTypeName(lpszValue)
End If
End Function

'Private Decode Resourcename name
Private Function DecodeResNameName(ByVal lpszValue As Long) As String
Dim Buffer() As Byte
Dim nLen As Long

If HiWord(lpszValue) Then
    nLen = lstrlenA(ByVal lpszValue)
    ReDim Buffer(0 To (nLen - 1)) As Byte
    CopyMemory Buffer(0), ByVal lpszValue, nLen
    DecodeResNameName = StrConv(Buffer, vbUnicode)
Else
    DecodeResNameName = CStr(lpszValue)
End If
End Function

'Public utility
Public Property Get HiWord(LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Property

'Public utility
Public Property Let HiWord(LongIn As Long, ByVal NewWord As Integer)
   Call CopyMemory(ByVal (VarPtr(LongIn) + 2), NewWord, 2)
End Property

'Public load the resourcelist
Public Sub EnumResData(Filename As String)
ResMod = LoadLibrary(Filename)
FrmMain.tView.Nodes.Clear
EnumResourceTypes ResMod, AddressOf CallBack_EnumResourceTypes, 0
FreeLibrary ResMod
End Sub
