''Imports VB = Microsoft.VisualBasic
Imports System.Data.OleDb
Imports System.IO

Module Module1

    Sub Main()
        Dim antw As String
        Dim prod As Boolean
        Dim steek As Boolean = False
        prod = False
        Console.WriteLine("databases van productie uitlezen ? ")
        antw = Console.ReadLine()
        If antw.Equals("Ja") Then
            prod = True
        End If
        Console.WriteLine("Steekproef dump ?")
        antw = Console.ReadLine()
        If antw.Equals("Ja") Then
            ''Dim connect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\temp\SteekProef.mdb"
            Dim connect As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\Steekproef.mdb;User Id=admin;Password=;"
            If prod Then
                connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\Uitvalteam AA\data\SteekProef.mdb;User Id=admin;Password=;"
            End If
            '' databasebestanden op c:\temp de externe koppelingen goed zetten !!!
            Using connectie1 As New OleDb.OleDbConnection(connect)
                Dim tabel As String = "Kantoren"
                Dim sqlStatement As String = String.Concat("SELECT * from ", tabel) '' is hier een dummy statement
				Try 
					connectie1.Open()
					Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(sqlStatement, connectie1)
					Dim oledbReader As OleDbDataReader = command.ExecuteReader()
					Dim table As DataTable = connectie1.GetSchema("Tables")
					connectie1.Close()
					Dim pad As String = "C:\temp\steekproef\"
					If prod Then
						pad = "C:\temp\prod\steekproef\"
					End If
					' Display the contents of the table.
					DisplayData(table, pad, connect)
				Catch ex As Exception
					Console.WriteLine(ex.Message)
				End	Try
            End Using
        End If
        Console.WriteLine("uitvalteam dump ?")
        antw = Console.ReadLine()
        If antw.Equals("Ja") Then
            Dim connect As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\uitvalteam.mdb;User Id=admin;Password=;"
            If prod Then
                connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\Uitvalteam AA\data\uitvalteam.mdb;User Id=admin;Password=;"
            End If
            '' databasebestanden op c:\temp de externe koppelingen goed zetten !!!
            Using connectie1 As New OleDb.OleDbConnection(connect)
                Dim tabel As String = "Aanbieders"
                Dim sqlStatement As String = String.Concat("SELECT * from ", tabel) '' is hier een dummy statement
                Try
                    connectie1.Open()
                    Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(sqlStatement, connectie1)
                    Dim oledbReader As OleDbDataReader = command.ExecuteReader()
                    Dim table As DataTable = connectie1.GetSchema("Tables")
                    connectie1.Close()
                    Dim pad As String = "C:\temp\uitvalteam\"
                    If prod Then
                        pad = "C:\temp\prod\uitvalteam\"
                    End If
                    ' Display the contents of the table.
                    DisplayData(table, pad, connect)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End Using
        End If
        Console.WriteLine("uitvalteamsoftware dump ?")
        antw = Console.ReadLine()
        If antw.Equals("Ja") Then
            Dim connect As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\temp\UitvalTeamSoftware2016.mdb;User Id=admin;Password=;"
            If prod Then
                connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\Uitvalteam AA\data\UitvalTeamSoftware2016.mdb;User Id=admin;Password=;"
            End If
            '' databasebestanden op c:\temp de externe koppelingen goed zetten !!!
            Using connectie1 As New OleDb.OleDbConnection(connect)
                Dim tabel As String = "Bonusrapport"
                Dim sqlStatement As String = String.Concat("SELECT * from ", tabel) '' is hier een dummy statement
                Try
                    connectie1.Open()
                    Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(sqlStatement, connectie1)
                    Dim oledbReader As OleDbDataReader = command.ExecuteReader()
                    Dim table As DataTable = connectie1.GetSchema("Tables")
                    connectie1.Close()
                    Dim pad As String = "C:\temp\uitvalsoft\"
                    If prod Then
                        pad = "C:\temp\prod\uitvalsoft\"
                    End If
                    ' Display the contents of the table.
                    DisplayData(table, pad, connect)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End Using
        End If
		Console.WriteLine("Press any key to continue.")
        Console.ReadKey()
    End Sub
    Private Sub DisplayData(ByVal table As DataTable, ByVal doelpad As String, ByVal connaam As String)
        Dim naam As String = ""
        For Each row As DataRow In table.Rows
            For Each col As DataColumn In table.Columns
                ''Console.WriteLine("{0} = {1}", col.ColumnName, row(col))
                If col.ColumnName.Equals("TABLE_NAME") Then
                    naam = row(col)
                    '' Console.WriteLine("{0} = {1}", naam, row(col))
                End If
                If col.ColumnName.Equals("TABLE_TYPE") Then
                    If row(col).Equals("TABLE") Then
                        Console.WriteLine("{0}", naam)
                        '' hier de subdumptable aanroepen met boven gevonden tabel naam
                        dumpTabelsteek(naam, doelpad, connaam)
                    End If
                End If
            Next
            Console.WriteLine("============================")
        Next
    End Sub
    ''' <summary>
    ''' '
    ''' </summary>
    ''' <param name="tabel"></param>
    ''' <param name="doel"></param>
    ''' <param name="connector"></param>
    Sub dumpTabelsteek(ByVal tabel As String, ByVal doel As String, ByVal connector As String)
        '' Dim ConnectieSteekproef As String
        'Dim da As OleDbDataAdapter
        'Dim rs As New DataSet
        Dim connectie As OleDb.OleDbConnection = New OleDb.OleDbConnection(connector)
        Try
            connectie.Open()
            Dim bestnaam As String = String.Concat(doel, tabel, ".txt")
            Dim sqlStatement As String
            sqlStatement = String.Concat("SELECT * from [", tabel, "]")
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(sqlStatement, connectie)
            Dim oledbReader As OleDbDataReader = command.ExecuteReader()
            Dim oWrite As StreamWriter
            oWrite = System.IO.File.CreateText(bestnaam)
            Console.WriteLine(String.Concat("best : ", bestnaam))
            Dim q As Char
            q = Chr(34)
            Dim schemaTable As DataTable
            Dim row As DataRow
            Dim column As DataColumn
            Dim count As Integer = 0
            Dim last As Integer
            Dim regel As String
            Dim kregel As String
            Dim velden As ArrayList = New ArrayList()
            Dim dubbel As String
            schemaTable = oledbReader.GetSchemaTable()
            kregel = ""
            regel = ""
            For Each row In schemaTable.Rows
                count = count + 1
                For Each column In schemaTable.Columns
                    'Console.WriteLine(String.Concat(column.ColumnName, " is ", row(column)))
                    If column.ColumnName Is "ColumnName" Then
                        '           Console.WriteLine(row(column))
                        kregel = String.Concat(kregel, row(column), ";")
                    End If
                    If column.ColumnName Is "DataType" Then
                        '            Console.WriteLine(row(column))
                        regel = String.Concat(regel, row(column), ";")
                        velden.Add(String.Concat(row(column)))
                    End If
                Next
            Next
            last = count - 1
            ' Console.WriteLine(String.Concat("aantal velden is ", count))
            'Console.WriteLine(kregel)
            'Console.WriteLine(regel)
            'Console.ReadLine()
            '' hier kopregel schrijven '' nog laatste ; van kop verwijderen ?
            oWrite.WriteLine(String.Concat(kregel))
            '' veld types ophalen uit array velden en aantal uit count
            While oledbReader.Read
                For iTel As Integer = 0 To count - 2
                    'Console.WriteLine(String.Concat(":", velden(iTel)))
                    If String.Compare(velden(iTel), "System.String", True) = 0 Then
                        'Console.WriteLine(String.Concat("String is ", oledbReader.Item(iTel)))
                        oWrite.Write(String.Concat(q, oledbReader.Item(iTel), q, ";"))
                    End If
                    '' hier dus andere veld types , zoals integer , boolean en double ook date-time
                    If String.Compare(velden(iTel), "System.Int32", True) = 0 Then
                        oWrite.Write(String.Concat(oledbReader.Item(iTel), ";"))
                    End If
                    If String.Compare(velden(iTel), "System.Int16", True) = 0 Then
                        oWrite.Write(String.Concat(oledbReader.Item(iTel), ";"))
                    End If
                    If String.Compare(velden(iTel), "System.Double", True) = 0 Then
                        '' verwissel de komma in de double voor een punt
                        dubbel = String.Concat(oledbReader.Item(iTel))
                        dubbel = Replace(dubbel, ",", ".")
                        oWrite.Write(String.Concat(dubbel, ";"))
                    End If
                    If String.Compare(velden(iTel), "System.Boolean", True) = 0 Then
                        oWrite.Write(String.Concat(oledbReader.Item(iTel), ";"))
                    End If
                    If String.Compare(velden(iTel), "System.DateTime", True) = 0 Then
                        oWrite.Write(String.Concat(oledbReader.Item(iTel), ";"))
                    End If


                Next
                '' hier een laastste veld wegschrijven zonder eind ";"
                If String.Compare(velden(last), "System.String", True) = 0 Then
                    'Console.WriteLine(String.Concat("String is ", oledbReader.Item(last)))
                    oWrite.Write(String.Concat(q, oledbReader.Item(last), q))
                End If
                '' hier dus andere veld types , zoals integer , boolean en double 
                If String.Compare(velden(last), "System.Int32", True) = 0 Then
                    oWrite.Write(String.Concat(oledbReader.Item(last)))
                End If
                If String.Compare(velden(last), "System.Int16", True) = 0 Then
                    oWrite.Write(String.Concat(oledbReader.Item(last)))
                End If
                If String.Compare(velden(last), "System.Double", True) = 0 Then
                    '' verwissel de komma in de double voor een punt
                    dubbel = String.Concat(oledbReader.Item(last))
                    dubbel = Replace(dubbel, ",", ".")
                    oWrite.Write(String.Concat(dubbel, ";"))
                End If
                If String.Compare(velden(last), "System.Boolean", True) = 0 Then
                    oWrite.Write(String.Concat(oledbReader.Item(last)))
                End If
                If String.Compare(velden(last), "System.DateTime", True) = 0 Then
                    oWrite.Write(String.Concat(oledbReader.Item(last)))
                End If

                'Console.ReadLine()
                oWrite.WriteLine()
                'oWrite.WriteLine(String.Concat(q, oledbReader.Item(0), q, ";", q, oledbReader.Item(1), q))
            End While
            'Console.ReadLine()
            oledbReader.Close()
            command.Dispose()
            connectie.Close()
            oWrite.Flush()
            oWrite.Close()
        Catch ex As Exception
            Throw New Exception("Fout bij dumptabel", ex)
        Finally
            connectie.Close()
        End Try

    End Sub

End Module
