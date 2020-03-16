Attribute VB_Name = "Module1"
Option Explicit

Public Type PlayerType
    playername As String * 15
    score As Integer
End Type

Public CurrentPlayer As PlayerType
Public Array_of_players() As PlayerType


Public NumberOfRecords As Integer
