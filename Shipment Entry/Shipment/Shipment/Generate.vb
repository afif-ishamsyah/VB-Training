﻿Imports System.Data.OleDb
Imports System.IO
Imports ACCPAC.Advantage

Public Class Generate

    Dim dt As New DataTable
    Dim filename As String
    Dim Username As String
    Dim Password As String
    Dim Database As String

    Public Sub LoadCsv()
        Dim folder = "C:\Users\SupportIT\Documents\Order Entry\Ongoing"
        Dim CnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & folder & ";Extended Properties=""text;HDR=Yes;FMT=Delimited(,)"";"

        Try
            Dim fileEntries As String() = Directory.GetFiles("C:\Users\SupportIT\Documents\Order Entry\Ongoing")
            filename = Path.GetFileName(fileEntries(0)) 'Get first file in the folder
            Using Adp As New OleDbDataAdapter("select * from [" + filename + "]", CnStr)
                Adp.Fill(dt) 'Insert CSV data to Datatable, CSV header excluded because in CnStr, HDR=Yes
                SendtoSage() 'Insert Datatble data to Sage
                dt.Clear()   'Clear Datatable
            End Using
            'If Insert to Sage success, move file Completed Folder
            File.Move("C:\Users\SupportIT\Documents\Order Entry\Ongoing\" + filename, "C:\Users\SupportIT\Documents\Order Entry\Completed\" + filename)
        Catch e As OleDbException 'Handle connection error in SendtoSage()
        Catch e As FileNotFoundException 'Handle when Folder is empty
        Catch e As IndexOutOfRangeException 'Handle when filename is already moved because of an error in Sendtosage() (look at SendtoSage exception)
        End Try

        Form1.Close()
    End Sub

    Public Sub SendtoSage()
        Dim session As Session
        Dim mDBLinkCmpRW As DBLink

        LoadFile()

        'Create new session
        session = New Session()
        session.Init("", "XY", "XY1000", "63A") 'first 3 parameter is always like that i dont know why, 4th parameter is Sage Version
        session.Open(Username, Password, Database, DateTime.Today, 0) 'Password and Username must be in UPPERCASE
        mDBLinkCmpRW = session.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)
        Try
            Dim OESHI1header As View
            Dim OESHI1detail1 As View
            Dim OESHI1detail2 As View
            Dim OESHI1detail3 As View
            Dim OESHI1detail4 As View
            Dim OESHI1detail5 As View
            Dim OESHI1detail6 As View
            Dim OESHI1detail7 As View
            Dim OESHI1detail8 As View
            Dim OESHI1detail9 As View
            Dim OESHI1detail10 As View
            Dim OESHI1detail11 As View
            Dim OESHI1detail12 As View


            'Open Shipment View, Look at Sage U.I.-------------------------------------------------------------------------------------

            OESHI1header = mDBLinkCmpRW.OpenView("OE0692")
            OESHI1detail1 = mDBLinkCmpRW.OpenView("OE0691")
            OESHI1detail2 = mDBLinkCmpRW.OpenView("OE0745")
            OESHI1detail3 = mDBLinkCmpRW.OpenView("OE0190")
            OESHI1detail4 = mDBLinkCmpRW.OpenView("OE0694")
            OESHI1detail5 = mDBLinkCmpRW.OpenView("OE0704")
            OESHI1detail6 = mDBLinkCmpRW.OpenView("OE0708")
            OESHI1detail7 = mDBLinkCmpRW.OpenView("OE0709")
            OESHI1detail8 = mDBLinkCmpRW.OpenView("OE0702")
            OESHI1detail9 = mDBLinkCmpRW.OpenView("OE0703")
            OESHI1detail10 = mDBLinkCmpRW.OpenView("OE0706")
            OESHI1detail11 = mDBLinkCmpRW.OpenView("OE0707")
            OESHI1detail12 = mDBLinkCmpRW.OpenView("OE0705")



            'Compose Shipment View, Look at Sage Macro when you create a Shipment-------------------------------------------------
            OESHI1header.Compose({OESHI1detail1, Nothing, OESHI1detail3, OESHI1detail2, OESHI1detail4, OESHI1detail5})
            OESHI1detail1.Compose({OESHI1header, Nothing, OESHI1detail8, OESHI1detail12, OESHI1detail9, OESHI1detail7, OESHI1detail6})
            OESHI1detail2.Compose({OESHI1header})
            OESHI1detail3.Compose({OESHI1header, OESHI1detail1})
            OESHI1detail4.Compose({OESHI1header})
            OESHI1detail5.Compose({OESHI1header})
            OESHI1detail6.Compose({OESHI1detail1, Nothing})
            OESHI1detail7.Compose({OESHI1detail1, Nothing})
            OESHI1detail8.Compose({OESHI1detail1})
            OESHI1detail9.Compose({OESHI1detail1, OESHI1detail10, Nothing, OESHI1detail11})
            OESHI1detail10.Compose({OESHI1detail9, Nothing})
            OESHI1detail11.Compose({OESHI1detail9, Nothing})
            OESHI1detail12.Compose({OESHI1detail1})

            'INSERT DATA TO ACCPAC

            OESHI1header.Init()
            OESHI1header.Fields.FieldByName("SHINUMBER").SetValue(dt.Rows(0)(0), False)
            OESHI1header.Fields.FieldByName("CUSTOMER").SetValue(dt.Rows(0)(5), False)
            OESHI1header.Fields.FieldByName("PONUMBER").SetValue(dt.Rows(0)(2), False)
            OESHI1header.Fields.FieldByName("SHIDATE").SetValue(Date.ParseExact(dt.Rows(0)(1).ToString(), "yyyyMMdd", Nothing), False)
            OESHI1header.Fields.FieldByName("DESC").SetValue(dt.Rows(0)(3), False)
            OESHI1header.Fields.FieldByName("REFERENCE").SetValue(dt.Rows(0)(4), False)

            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    OESHI1detail1.RecordCreate(ViewRecordCreate.NoInsert)
                    OESHI1detail1.Fields.FieldByName("ITEM").SetValue(dt.Rows(i)(7), False)
                    OESHI1detail1.Fields.FieldByName("LOCATION").SetValue(dt.Rows(i)(8), False)
                    OESHI1detail1.Fields.FieldByName("QTYSHIPPED").SetValue(dt.Rows(i)(9), False)
                    OESHI1detail1.Fields.FieldByName("PRIUNTPRC").SetValue(dt.Rows(i)(11), False)
                    OESHI1detail1.Fields.FieldByName("SHIUNIT").SetValue(dt.Rows(i)(10), False)
                    OESHI1detail1.Insert()
                Next
            End If

            OESHI1header.Insert()

            'Handle error is Datatable fail to be inserted to Sage, move the errored CSV file, and create an error log 
        Catch e As Runtime.InteropServices.COMException
            Dim errors As List(Of String) = New List(Of String)
            Dim files As FileStream = File.Create("C:\Users\SupportIT\Documents\Order Entry\ErrorLog\" + filename + ".txt")
            files.Close()

            For k As Integer = 0 To session.Errors.Count() - 1
                errors.Add(session.Errors(k).Message)
            Next

            Dim errorMessage As String = String.Join(" ", errors)


            My.Computer.FileSystem.WriteAllText("C:\Users\SupportIT\Documents\Order Entry\ErrorLog\" + filename + ".txt", errorMessage, True)
            File.Move("C:\Users\SupportIT\Documents\Order Entry\Ongoing\" + filename, "C:\Users\SupportIT\Documents\Order Entry\ErrorFile\" + filename)
            session.Errors.Clear()


        End Try
    End Sub

    Public Sub LoadFile()

        Dim fileload As String = "C:\Users\SupportIT\Documents\C# Training\Database Setup\Database Setup\DatabaseSetup.txt"
        Dim lines() As String
        Dim loadedLines() As String = File.ReadAllLines(fileload)

        Dim index As Integer = 0

        Dim n As Integer = Integer.Parse(loadedLines(index))
        lines = New String(n) {}
        Array.Copy(loadedLines, (index + 1), lines, 0, n)
        Username = lines(n - 1)

        index = (index + 2)
        n = Integer.Parse(loadedLines(index))
        lines = New String(n) {}
        Array.Copy(loadedLines, (index + 1), lines, 0, n)
        Password = lines(n - 1)

        index = (index + 2)
        n = Integer.Parse(loadedLines(index))
        lines = New String(n) {}
        Array.Copy(loadedLines, (index + 1), lines, 0, n)
        Database = lines(n - 1)



    End Sub
End Class