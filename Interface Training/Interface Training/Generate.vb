Imports System.Data.OleDb
Imports System.IO
Imports ACCPAC.Advantage

Public Class Generate

    Dim dt As New DataTable
    Dim filename As String

    Public Sub LoadCsv()
        Dim folder = "C:\Users\SupportIT\Documents\Purchase Order\Ongoing"
        Dim CnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & folder & ";Extended Properties=""text;HDR=Yes;FMT=Delimited(,)"";"

        Try
            Dim fileEntries As String() = Directory.GetFiles("C:\Users\SupportIT\Documents\Purchase Order\Ongoing")
            filename = Path.GetFileName(fileEntries(0)) 'Get first file in the folder
            Using Adp As New OleDbDataAdapter("select * from [" + filename + "]", CnStr)
                Adp.Fill(dt) 'Insert CSV data to Datatable, CSV header excluded because in CnStr, HDR=Yes
                SendtoSage() 'Insert Datatble data to Sage
                dt.Clear()   'Clear Datatable
            End Using
            'If Insert to Sage success, move file Completed Folder
            File.Move("C:\Users\SupportIT\Documents\Purchase Order\Ongoing\" + filename, "C:\Users\SupportIT\Documents\Purchase Order\Completed\" + filename)
        Catch e As OleDbException 'Handle connection error in SendtoSage()
        Catch e As FileNotFoundException 'Handle when Folder is empty
        Catch e As IndexOutOfRangeException 'Handle when filename is already moved because of an error in Sendtosage() (look at SendtoSage exception)
        End Try

        Form1.Close()
    End Sub

    Public Sub SendtoSage()
        Dim session As Session
        Dim mDBLinkCmpRW As DBLink

        'Create new session
        session = New Session()
        session.Init("", "XY", "XY1000", "63A") 'first 3 parameter is always like that i dont know why, 4th parameter is Sage Version
        session.Open("ADMIN", "SUP3RVIS0RCMW", "CMWTRN", DateTime.Today, 0) 'Password and Username must be in UPPERCASE
        mDBLinkCmpRW = session.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)
        Try
            Dim POPOR1header As View
            Dim POPOR1detail1 As View
            Dim POPOR1detail2 As View
            Dim POPOR1detail3 As View
            Dim POPOR1detail4 As View
            Dim POPOR1detail5 As View
            Dim POPOR1detail6 As View

            'Open PO View, Look at Sage U.I.-------------------------------------------------------------------------------------
            POPOR1header = mDBLinkCmpRW.OpenView("PO0620")
            POPOR1detail1 = mDBLinkCmpRW.OpenView("PO0630")
            POPOR1detail2 = mDBLinkCmpRW.OpenView("PO0610")
            POPOR1detail3 = mDBLinkCmpRW.OpenView("PO0632")
            POPOR1detail4 = mDBLinkCmpRW.OpenView("PO0619")
            POPOR1detail5 = mDBLinkCmpRW.OpenView("PO0623")
            POPOR1detail6 = mDBLinkCmpRW.OpenView("PO0633")

            'Compose PO View, Look at Sage Macro when you create a Purchase Order-------------------------------------------------
            POPOR1header.Compose({POPOR1detail2, POPOR1detail1, POPOR1detail3, POPOR1detail4, POPOR1detail5})
            POPOR1detail1.Compose({POPOR1header, POPOR1detail2, POPOR1detail4, POPOR1detail6})
            POPOR1detail2.Compose({POPOR1header, POPOR1detail1})
            POPOR1detail3.Compose({POPOR1header, POPOR1detail4})
            POPOR1detail4.Compose({POPOR1header, POPOR1detail2, POPOR1detail1, POPOR1detail3})
            POPOR1detail5.Compose({POPOR1header})
            POPOR1detail6.Compose({POPOR1detail1})

            'INSERT DATA TO ACCPAC

            'Create Header------------------------------------------------------------------------------------------------
            'Only create, will be added to Sage after we create the Optional Fields
            POPOR1header.Init()
            POPOR1header.Fields.FieldByName("PONUMBER").SetValue(dt.Rows(0)(0), False)
            POPOR1header.Fields.FieldByName("VDCODE").SetValue(dt.Rows(0)(5), False)
            POPOR1header.Fields.FieldByName("DATE").SetValue(Date.ParseExact(dt.Rows(0)(1).ToString(), "yyyyMMdd", Nothing), False)
            POPOR1header.Fields.FieldByName("FOBPOINT").SetValue(dt.Rows(0)(8), False)
            POPOR1header.Fields.FieldByName("EXPARRIVAL").SetValue(Date.ParseExact(dt.Rows(0)(2).ToString(), "yyyyMMdd", Nothing), False)
            POPOR1header.Fields.FieldByName("STCODE").SetValue(dt.Rows(0)(12), False)
            If dt.Rows(0)(3).ToString() = "undefined" Then
                POPOR1header.Fields.FieldByName("DESCRIPTIO").SetValue("", False)
            Else
                POPOR1header.Fields.FieldByName("DESCRIPTIO").SetValue(dt.Rows(0)(3), False)
            End If
            POPOR1header.Fields.FieldByName("REFERENCE").SetValue(dt.Rows(0)(4), False)
            '---------------------------------------------------------------------------------------------------------------

            'Create and Insert Line Item--------------------------------------------------------------------------------------
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    POPOR1detail1.RecordCreate(ViewRecordCreate.NoInsert)
                    If dt.Rows(i)(13).ToString() = "undefined" Then
                        POPOR1detail1.Fields.FieldByName("ITEMNO").SetValue("", False)
                    Else
                        POPOR1detail1.Fields.FieldByName("ITEMNO").SetValue(dt.Rows(i)(13), False)
                    End If
                    If dt.Rows(i)(14).ToString() = "undefined" Then
                        POPOR1header.Fields.FieldByName("ITEMDESC").SetValue("", False)
                    Else
                        POPOR1detail1.Fields.FieldByName("ITEMDESC").SetValue(dt.Rows(i)(14), False)
                    End If

                    POPOR1detail1.Fields.FieldByName("LOCATION").SetValue(dt.Rows(i)(21), False)
                    POPOR1detail1.Fields.FieldByName("OQORDERED").SetValue(dt.Rows(i)(15), False)
                    POPOR1detail1.Fields.FieldByName("ORDERUNIT").SetValue(dt.Rows(i)(16), False)
                    POPOR1detail1.Fields.FieldByName("UNITCOST").SetValue(dt.Rows(i)(17), False)
                    POPOR1detail1.Fields.FieldByName("EXTENDED").SetValue(dt.Rows(i)(18), False)
                    POPOR1detail1.Fields.FieldByName("EXPARRIVAL").SetValue(Date.ParseExact(dt.Rows(0)(2).ToString(), "yyyyMMdd", Nothing), False)
                    POPOR1detail1.Fields.FieldByName("GLACEXPENS").SetValue(dt.Rows(i)(19), False)

                    'Create and Insert Comment, Because every 60 characters, comment must be added to a new line
                    POPOR1detail2.Init()
                    If dt.Rows(i)(7).ToString() = "undefined" Then
                        POPOR1detail1.Fields.FieldByName("HASCOMMENT").SetValue("0", False)
                    ElseIf dt.Rows(i)(7).ToString().Length > 60 Then
                        POPOR1detail1.Fields.FieldByName("HASCOMMENT").SetValue("1", False)
                        Dim length As Integer = dt.Rows(i)(7).ToString().Length
                        Dim loops As Integer = length \ 60
                        Dim mods As Integer = length Mod 60
                        For j As Integer = 0 To loops - 1
                            POPOR1detail2.Fields.FieldByName("COMMENT").SetValue(dt.Rows(i)(7).ToString().Substring(60 * j, 60), False)
                            POPOR1detail2.Insert()
                        Next
                        POPOR1detail2.Fields.FieldByName("COMMENT").SetValue(dt.Rows(i)(7).ToString().Substring(60 * loops, mods), False)
                        POPOR1detail2.Insert()

                    ElseIf dt.Rows(i)(7).ToString().Length <= 60 Then
                        POPOR1detail1.Fields.FieldByName("HASCOMMENT").SetValue("1", False)
                        POPOR1detail2.Fields.FieldByName("COMMENT").SetValue(dt.Rows(i)(7).ToString(), False)
                        POPOR1detail2.Insert()
                    End If
                    '-------------------------------------------------------------------------------------------------------------

                    POPOR1detail1.Insert()
                Next
            End If

            'Create and Insert Optional Fields------------------------------------------------------------------------------
            POPOR1detail5.Fields.FieldByName("OPTFIELD").SetValue("PRDATE", False)
            POPOR1detail5.Fields.FieldByName("VALIFDATE").SetValue(Date.ParseExact(dt.Rows(0)(22).ToString(), "yyyyMMdd", Nothing), False)
            POPOR1detail5.Insert()
            POPOR1detail5.Fields.FieldByName("OPTFIELD").SetValue("PRDATECOM", False)
            POPOR1detail5.Fields.FieldByName("VALIFDATE").SetValue(Date.ParseExact(dt.Rows(0)(23).ToString(), "yyyyMMdd", Nothing), False)
            POPOR1detail5.Insert()
            POPOR1detail5.Fields.FieldByName("OPTFIELD").SetValue("PRNO", False)
            POPOR1detail5.Fields.FieldByName("VALIFTEXT").SetValue(dt.Rows(0)(24), False)
            POPOR1detail5.Insert()
            '----------------------------------------------------------------------------------------------------------------

            'Insert Header
            POPOR1header.Insert()

            'Handle error is Datatable fail to be inserted to Sage, move the errored CSV file, and create an error log 
        Catch e As Runtime.InteropServices.COMException
            Dim errors As List(Of String) = New List(Of String)
            Dim files As FileStream = File.Create("C:\Users\SupportIT\Documents\Purchase Order\ErrorLog\" + filename + ".txt")
            files.Close()

            For k As Integer = 0 To session.Errors.Count() - 1
                If session.Errors(k).Message = "Tax group  does not exist." Then
                    errors.Add("Vendor not exist.")
                End If
                errors.Add(session.Errors(k).Message)
            Next

            Dim errorMessage As String = String.Join(" ", errors)


            My.Computer.FileSystem.WriteAllText("C:\Users\SupportIT\Documents\Purchase Order\ErrorLog\" + filename + ".txt", errorMessage, True)
            File.Move("C:\Users\SupportIT\Documents\Purchase Order\Ongoing\" + filename, "C:\Users\SupportIT\Documents\Purchase Order\ErrorFile\" + filename)
            session.Errors.Clear()


        End Try
    End Sub
End Class
