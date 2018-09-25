VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBTableLinkHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
' Private methods
'
Private Function getDBPathFromTablePath(ByRef dbPath As String) As String
   Dim pos As Long

   pos = InStr(1, dbPath, "=")
   
   If pos <> 0 Then
      getDBPathFromTablePath = Right$(dbPath, Len(dbPath) - pos)
   Else
      getDBPathFromTablePath = CurrentProject.Path
   End If
End Function

'
' Public methods
'
Public Function getDBPath(ByRef tableName As String) As String
   Dim db As DAO.Database

   Set db = DBEngine.Workspaces(0).Databases(0)
   
   Dim tdf As DAO.TableDef
   
   Set tdf = db.TableDefs(tableName)

   getDBPath = getDBPathFromTablePath(tdf.Connect)
   
   Set tdf = Nothing
   
   Set db = Nothing
End Function

Public Function changeDBPath(ByRef fromDBPath As String, ByRef toDBPath As String) As Integer
   Dim count As Integer
   
   Dim db As DAO.Database

   Set db = DBEngine.Workspaces(0).Databases(0)
   
   Dim tdf As DAO.TableDef
   
   For Each tdf In db.TableDefs
      If Len(tdf.Connect) <> 0 Then
         If getDBPathFromTablePath(tdf.Connect) = fromDBPath Then
            tdf.Connect = ";DATABASE=" & toDBPath
            count = count + 1
         End If
      End If
   Next

   changeDBPath = count

   db.TableDefs.Refresh

   Set db = Nothing
End Function
