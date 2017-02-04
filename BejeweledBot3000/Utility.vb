Public Module Utility
    Public TileCount As Integer = 8
    Public TileSize As Integer = 50

    Declare Function apimouse_event Lib "user32.dll" Alias "mouse_event" (ByVal dwFlags As Int32, ByVal dX As Int32, ByVal dY As Int32, ByVal cButtons As Int32, ByVal dwExtraInfo As Int32) As Boolean
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Declare Function GetWindowRect Lib "user32.dll" Alias "GetWindowRect" (hwnd As IntPtr, ByRef rectangle As Rect) As IntPtr
    Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal ByValnVirtKey As Int32) As Int16
    Public Const VK_CAPSLOCK = &H14

    Public Structure Rect
        Public Left As Int32
        Public Top As Int32
        Public Right As Int32
        Public Bottom As Int32
    End Structure
End Module
