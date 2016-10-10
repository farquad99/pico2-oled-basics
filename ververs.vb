Option Compare Database

'' verversen van tabellen van uit andere database
'' met tabel namen en database namen in tabel Z-Tables-TO-Update
Sub ververs()
Dim r As Recordset
Dim CurrentTableName As String
Set r = CurrentDb.OpenRecordset("Z_Tables_TO_Update")
r.MoveFirst
Do Until r.EOF
  CurrentTableName = r!TabelNaam
  If TableExists(CurrentTableName, True) Then
    DoCmd.DeleteObject acTable, CurrentTableName
  End If
  If Not TableExists(r!TabelNaam, True) Then
    DoCmd.TransferDatabase acImport, "Microsoft Access", r!DatabaseNaam, acTable, r!TabelNaam, r!TabelNaam
  End If
  r.MoveNext
Loop
End Sub
''
Function IsTable(sTblName As String) As Boolean
    'does table exists and work ?
    'note: finding the name in the TableDefs collection is not enough,
    '      since the backend might be invalid or missing

    On Error GoTo hello
    Dim x
    x = DCount("*", sTblName)
    IsTable = True
    Exit Function
hello:
    Debug.Print Now, sTblName, Err.Number, Err.Description
    IsTable = False

End Function
''
Public Function TableExists(strTableName As String, Optional ysnRefresh As Boolean, Optional db As DAO.Database) As Boolean
' Originally Based on Tony Toews function in TempTables.MDB, http://www.granite.ab.ca/access/temptables.htm
' Based on testing, when passed an existing database variable, this is the fastest
On Error GoTo errHandler
  Dim tdf As DAO.TableDef

  If db Is Nothing Then Set db = CurrentDb()
  If ysnRefresh Then db.TableDefs.Refresh
  Set tdf = db(strTableName)
  TableExists = True

exitRoutine:
  Set tdf = Nothing
  Exit Function

errHandler:
  Select Case Err.Number
    Case 3265
      TableExists = False
    Case Else
      MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error in mdlBackup.TableExists()"
  End Select
  Resume exitRoutine
End Function

'' nog uitzoeken wat Access.Application is
Public Function TableExistsO(theDatabase As Access.Application, _
    tableName As String) As Boolean

    ' Presume that table does not exist.
    TableExistsO = False

    ' Define iterator to query the object model.
    Dim iTable As Integer

    ' Loop through object catalogue and compare with search term.
    For iTable = 0 To theDatabase.CurrentData.AllTables.Count - 1
        If theDatabase.CurrentData.AllTables(iTable).Name = tableName Then
            TableExistsO = True
            Exit Function
        End If
    Next iTable

End Function


Public Function TableExists2(strTableName As String, Optional ysnRefresh As Boolean, Optional db As DAO.Database) As Boolean
On Error GoTo errHandler
  Dim bolCleanupDB As Boolean
  Dim tdf As DAO.TableDef

  If db Is Nothing Then
     Set db = CurrentDb()
     bolCleanupDB = True
  End If
  If ysnRefresh Then db.TableDefs.Refresh
  For Each tdf In db.TableDefs
    If tdf.Name = strTableName Then
       TableExists2 = True
       Exit For
    End If
  Next tdf

exitRoutine:
  Set tdf = Nothing
  If bolCleanupDB Then
     Set db = Nothing
  End If
  Exit Function

errHandler:
  MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error in mdlBackup.TableExists1()"
  Resume exitRoutine
End Function
