Attribute VB_Name = "modFormRuler"
Option Explicit

Public Type PointAPI
  X As Long
  Y As Long
End Type


Public Function TwipToInch(pTwip As Single) As String
    TwipToInch = Format(Val(pTwip) / 1440, "#.##") & " in"
End Function
