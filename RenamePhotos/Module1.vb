Imports System
Imports System.IO
Imports System.Data.Odbc
Imports System.Data.SqlClient
Module Module1

    Sub Main()
        Dim sourcePath As String = "C:\Admin\RenameFiles\Media"
        Dim searchPattern As String = "*.jpg"

        '-- read from table
        Console.WriteLine("Before")
        'ConnectToSQL()
        Console.WriteLine("After")

        '-- rename file
        Dim i As Integer = 1
        For Each fileName As String In Directory.GetFiles(sourcePath, searchPattern, SearchOption.AllDirectories)
            '--get LookupID 
            'Console.WriteLine(System.IO.Path.GetFileName(fileName))
            Console.WriteLine(GetIDfromFileName(System.IO.Path.GetFileName(fileName)))
            'Console.WriteLine(GetPhotoAddDateFromFileName(System.IO.Path.GetFileName(fileName)))
            ConnectToSQL2(fileName, GetIDfromFileName(System.IO.Path.GetFileName(fileName)), GetPhotoAddDateFromFileName(System.IO.Path.GetFileName(fileName)))

            'Console.WriteLine(fileName.IndexOf("-"))

            '--File.Move(Path.Combine(sourcePath, fileName), Path.Combine(sourcePath, fileName + "Stan_" & i & ".jpg"))
            ' i += 1
        Next
    End Sub

    Public Function GetIDfromFileName(ByVal filename As String) As String
        Dim LookupID As String
        LookupID = Left(filename, InStr(filename, "-") - 1)
        Return LookupID
    End Function

    Public Function GetPhotoAddDateFromFileName(ByVal filename As String) As String
        Dim PhotoAddDate As String
        PhotoAddDate = Left(Right(filename, 23), 10)
        Return PhotoAddDate
    End Function

    Public Function ConnectToSQL2(ByVal filename As String, ByVal LookupID As String, ByVal PhotoAddDate As String) As String
        Dim con As New SqlConnection
        Dim reader As SqlDataReader

        Dim sourcePath As String = "C:\Admin\RenameFiles\Media"
        Dim searchPattern As String = "*.jpg"

        Try
            con.ConnectionString = "Data Source=CTBBDEV01;Initial Catalog=STC-ODS-Dev;Integrated Security=SSPI;Persist Security Info=False;"
            '--Dim cmd As New SqlCommand("SELECT P.URL, P.AgeAtPhoto FROM Photos P join Children C ON C.ChildID=P.Child_ChildID where C.ChildLookupID='" + LookupID + "'", con)

            '--Dim cmd As New SqlCommand("select 'https://stcodsphotos.imgix.net/', case when dbo.UFN_ODS_GETCOUNTRYNAME(L.COUNTRYOFFICE_LOCATIONID) LIKE '%_USA%' then 'USA' when dbo.UFN_ODS_GETCOUNTRYNAME(L.COUNTRYOFFICE_LOCATIONID) LIKE '%Nepal%' then 'Nepal' else dbo.UFN_ODS_GETCOUNTRYNAME(L.COUNTRYOFFICE_LOCATIONID) end, C.FirstName, CONVERT(VARCHAR(10),P.PhotoAddDate,120), 'headshot', case when P.ISMAIN=1 then 'Y' when P.ISMAIN=0 then  'N' end AS ISMAIN FROM Photos P JOIN Children C ON C.ChildID=P.Child_ChildID join LOCATIONS L ON C.[Community_LocationID]=L.LOCATIONID where C.ChildLookupID='" + LookupID + "' and CONVERT(DATE,P.PhotoAddDate)='" + PhotoAddDate + "'", con)    '--'10100594' and CONVERT(DATE,P.PhotoAddDate)='2016-04-27' 
            Dim cmd As New SqlCommand("select top 1 'https://stcodsphotos.imgix.net/', P.PhotoType, C.FirstName, 'headshot', CONVERT(VARCHAR(10),P.PhotoAddDate,120), case when P.ISMAIN='True' then 'Y' when P.ISMAIN='False' then  'N' end  FROM Photos P JOIN Children C ON C.ChildID=P.Child_ChildID join LOCATIONS L ON C.[Community_LocationID]=L.LOCATIONID where C.ChildLookupID='" + LookupID + "' and CONVERT(DATE,P.PhotoAddDate)='" + PhotoAddDate + "'", con)    '--'10100594' and CONVERT(DATE,P.PhotoAddDate)='2016-04-27' 

            con.Open()
            Console.WriteLine("Connection Opened")


            ' Execute Query    '
            reader = cmd.ExecuteReader()
            While reader.Read()
                Console.WriteLine(String.Format("{0}, {1}, {2}, {3}, {4}, {5}", _
                   reader(0), reader(1), reader(2), reader(3), reader(4), reader(5)))

                File.Move(Path.Combine(sourcePath, filename), Path.Combine(sourcePath, reader(1) & "-" & reader(2) & "-" & LookupID & "-" & reader(3) & "-" & reader(4) & "-" & reader(5) & ".jpg"))
                'NOTE: (^^) You are trying to read 2 columns here, but you only        '
                '   SELECT-ed one (username) originally...                             '
                ' , Also, you can call the columns by name(string), not just by number '

            End While
        Catch ex As Exception
            Console.WriteLine("Error while connecting to SQL Server." & ex.Message)
        Finally
            con.Close() 'Whether there is error or not. Close the connection.    '
        End Try
        'Return reader { note: reader is not valid after it is closed }          '
        Console.WriteLine()
        Return "done"
    End Function

    Public Function ConnectToSQL() As String
        Dim con As New SqlConnection
        Dim reader As SqlDataReader
        Try
            con.ConnectionString = "Data Source=wes02308\SQL2014;Initial Catalog=BBEC_ODS;Integrated Security=SSPI;Persist Security Info=False;"
            Dim cmd As New SqlCommand("SELECT top 1000 URL, AgeAtPhoto FROM Photos", con)
            con.Open()
            Console.WriteLine("Connection Opened")


            ' Execute Query    '
            reader = cmd.ExecuteReader()
            While reader.Read()
                Console.WriteLine(String.Format("{0}, {1}", _
                   reader(0), reader(1)))
                'NOTE: (^^) You are trying to read 2 columns here, but you only        '
                '   SELECT-ed one (username) originally...                             '
                ' , Also, you can call the columns by name(string), not just by number '

            End While
        Catch ex As Exception
            Console.WriteLine("Error while connecting to SQL Server." & ex.Message)
        Finally
            con.Close() 'Whether there is error or not. Close the connection.    '
        End Try
        'Return reader { note: reader is not valid after it is closed }          '
        Console.WriteLine()
        Return "done"
    End Function

End Module
