Attribute VB_Name = "Module3"
Option Explicit
Type make_buffer
   Index As String
   number As String
   name As String
   ban As String
   che As String
   flag As String
   now_che As String
End Type
Global china_english As Integer
Global MakeBuffer(100) As make_buffer
Global run_flag As Integer
Global gbFindFailed As Integer
Global gsFindField As String
Global gsFindExpr As String
Global gsFindOp As String

