VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public x_byte1 As Long
Public x_byte2 As Long
Public y_byte1 As Long
Public y_byte2 As Long
Public site As Long
Public bin As Long
Public x_loc As Long
Public y_loc As Long
Public id As Long

Private Sub Class_Initialize()
    
    x_byte1 = 0
    x_byte2 = 0
    y_byte1 = 0
    y_byte2 = 0
    site = 0
    bin = 0
    x_loc = 0
    y_loc = 0
    id = 0

End Sub

Public Sub cal()

    x_loc = x_byte2 * 256 + x_byte1
    y_loc = y_byte2 * 256 + y_byte1

End Sub
