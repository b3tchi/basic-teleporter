Attribute VB_Name = "mod_PSClassCreator"
Option Compare Database

Public Function GetOptions() As Object

    Set GetOptions = modObjects.Options

End Function

Public Function GetVcsIndex() As Object

    Set GetVcsIndex = modObjects.VCSIndex

End Function


Public Function CloseVcsIndex()

    Set VCSIndex = Nothing

End Function
