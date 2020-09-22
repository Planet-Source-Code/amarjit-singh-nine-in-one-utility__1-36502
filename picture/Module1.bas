Attribute VB_Name = "Module1"
Option Explicit

Public Function getred(ByVal color As Long) As Long
getred = color Mod 256
End Function
Public Function getgreen(ByVal color As Long) As Long
getgreen = (color And &HFF00FF00) / (256)
End Function
Public Function getblue(ByVal color As Long) As Long
getblue = (color And &HFFFF0000) / (65536)
End Function

'Public Type sfont
'size As Integer
'cx As Integer
'cy As Integer
'End Type
