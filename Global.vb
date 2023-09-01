Imports System.Data.SQLite

Module GlobalCode
    ' Define a class to represent a Borehole
    Public Class BoreHole
        Public Id As Short        ' Unique identifier for the borehole
        Public SiteName As String ' Name of the site where the borehole is located
        Public Location As String ' Geographic location of the borehole
        Public Depth As Single    ' Depth of the borehole
        Public BaseFile As String ' File associated with the borehole (if any)
    End Class

    ' Declare variables for SQLite connection, command, data reader, and data adapter
    Dim sqlite_conn As SQLiteConnection
    Dim sqlite_cmd As SQLiteCommand
    Dim sqlite_datareader As SQLiteDataReader
    Dim sqliteAdapter As SQLiteDataAdapter

    ' Open the SQLite database connection
    Sub OpenDatabase()
        ' Create a new database connection: with file data.sqlite
        sqlite_conn = New SQLiteConnection("Data Source=" & Application.LocalUserAppDataPath & "\data.sqlite;Version=3;")

        ' Open the connection
        sqlite_conn.Open()

        ' Create a new SQLite command
        sqlite_cmd = sqlite_conn.CreateCommand()

        ' Create the 'Boreholes' table if it doesn't exist
        sqlite_cmd.CommandText =
           "CREATE TABLE IF NOT EXISTS
              [Boreholes] (
              [Id]       INTEGER NOT NULL PRIMARY KEY,
              [SITENAME] VARCHAR(256) NULL,
              [LOCATION] VARCHAR(256) NULL,
              [DEPTH]    DOUBLE(10,4) NOT NULL,
              [BASEFILE] VARCHAR(256))"


        ' Execute the SQL command
        sqlite_cmd.ExecuteNonQuery()
    End Sub

    ' Close the SQLite database connection
    Sub CloseDatabase()
        ' Dispose of the command and close the connection
        sqlite_cmd.Dispose()
        sqlite_conn.Close()
    End Sub

    ' Add a new Borehole record to the database
    Function AddBorehole(ByRef bh As BoreHole) As Boolean
        Dim result As Short
        sqlite_cmd.CommandText =
            " INSERT INTO Boreholes (
                [Id], [SITENAME], [LOCATION], [DEPTH], [BASEFILE] )
              VALUES (@ID, @SiteName, @Location, @Depth, '')"
        sqlite_cmd.Parameters.AddWithValue("@ID", bh.Id)
        sqlite_cmd.Parameters.AddWithValue("@SiteName", bh.SiteName)
        sqlite_cmd.Parameters.AddWithValue("@Location", bh.Location)
        sqlite_cmd.Parameters.AddWithValue("@Depth", bh.Depth)
        Try
            ' Execute the SQL command and return success status
            result = sqlite_cmd.ExecuteNonQuery()
        Catch
            Return False
        End Try
        Return True
    End Function
    'test

    ' Update an existing Borehole record in the database
    Function UpdateBorehole(ByRef bh As BoreHole) As Boolean
        Dim result As Short, bnAddBaseFile As Boolean = True
        If bh.BaseFile.Length < 2 Then bnAddBaseFile = False

        If bnAddBaseFile Then
            ' Update record including BaseFile
            sqlite_cmd.CommandText =
            " UPDATE Boreholes SET [SITENAME]=@SiteName, [LOCATION]=@Location, [DEPTH]=@Depth, [BASEFILE]=@BaseFile 
              WHERE [Id]=@ID"
        Else
            ' Update record without BaseFile
            sqlite_cmd.CommandText =
            " UPDATE Boreholes SET [SITENAME]=@SiteName, [LOCATION]=@Location, [DEPTH]=@Depth WHERE [Id]=@ID"
        End If
        sqlite_cmd.Parameters.AddWithValue("@ID", bh.Id)
        sqlite_cmd.Parameters.AddWithValue("@SiteName", bh.SiteName)
        sqlite_cmd.Parameters.AddWithValue("@Location", bh.Location)
        sqlite_cmd.Parameters.AddWithValue("@Depth", bh.Depth)
        If bnAddBaseFile Then
            sqlite_cmd.Parameters.AddWithValue("@BaseFile", bh.BaseFile)
        End If
        Try
            ' Execute the SQL command and return success status
            result = sqlite_cmd.ExecuteNonQuery()
        Catch
            Return False
        End Try
        Return True
    End Function

    ' Delete a Borehole record from the database based on its ID
    Function DeleteBorehole(ByRef id As Short) As Short
        sqlite_cmd.CommandText = " DELETE FROM Boreholes WHERE Id=" & id
        Return sqlite_cmd.ExecuteNonQuery()
    End Function

    ' Delete all Borehole records from the database
    Function _DeleteAllBoreholes() As Short
        sqlite_cmd.CommandText = " DELETE FROM Boreholes"
        Return sqlite_cmd.ExecuteNonQuery()
    End Function

    ' Retrieve a list of Borehole objects from the database
    Function GetBoreholes() As List(Of BoreHole)
        Dim bh As New List(Of BoreHole)
        sqlite_cmd.CommandText = "SELECT Id, SITENAME, LOCATION, DEPTH, BASEFILE FROM Boreholes ORDER BY Id"

        sqlite_datareader = sqlite_cmd.ExecuteReader()

        ' Loop through the data reader and create Borehole objects
        Do While (sqlite_datareader.Read())
            bh.Add(New BoreHole With {.Id = sqlite_datareader.GetValue(0), .SiteName = sqlite_datareader.GetValue(1), .Location = sqlite_datareader.GetValue(2), .Depth = sqlite_datareader.GetValue(3), .BaseFile = "" & sqlite_datareader.GetValue(4)})
        Loop
        sqlite_datareader.Close()
        Return bh
    End Function

    ' Read data from a CSV file and return as a 2D string array
    Function ReadCSVFile(ByRef FileName As String) As String()()
        Dim data As New List(Of String())()

        Try
            Using MyReader As New FileIO.TextFieldParser(FileName)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")

                ' Read CSV data line by line and add to the data list
                While Not MyReader.EndOfData
                    Try
                        Dim split As String() = MyReader.ReadFields()
                        data.Add(split)
                    Catch ex As FileIO.MalformedLineException
                        ' Skip invalid lines
                        ReadCSVFile = Nothing
                    End Try
                End While
            End Using
            Return data.ToArray()
        Catch ex As System.Exception
            ' Display error message and return Nothing
            MsgBox(ex.Message, vbOKOnly Or vbExclamation, "File Read")
            Return Nothing
        End Try
    End Function

    ' Get the directory path for a Borehole based on its number
    Function GetBoreholeDirectory(ByRef bhnum As Short) As String
        Return Application.LocalUserAppDataPath & "\" & bhnum.ToString().PadLeft(2, "0")
    End Function

End Module
