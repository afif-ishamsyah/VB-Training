Imports System.IO
Imports ACCPAC.Advantage
Imports Microsoft.Office.Interop

Public Class FileLabel

    Dim FileExcel As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FileNameTextbox.ReadOnly = True
        UploadButton.Enabled = False
        LoadingLabel.Visible = False
    End Sub

    Private Sub UploadButton_Click(sender As Object, e As EventArgs) Handles UploadButton.Click
        Dim extension As String = Path.GetExtension(FileExcel)
        If extension = ".xlsx" Or extension = ".xls" Then
            LoadingLabel.Visible = True
            SenDatoSage(FileExcel)
        Else
            MsgBox("File must be a Microsoft Excel file (.xlsx or .xls)", 0, "Wrong File Type")
        End If

    End Sub

    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
        If SearchDialog.ShowDialog() = DialogResult.OK Then
            FileExcel = SearchDialog.FileName
            FileNameTextbox.Text = Path.GetFileName(FileExcel)
            If FileNameTextbox.Text <> "" Then
                UploadButton.Enabled = True
            End If
        End If

    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        End
    End Sub



    Private Sub SenDatoSage(Path As String)
        SearchButton.Enabled = False
        CancelButton.Enabled = False
        UploadButton.Enabled = False
        DatabaseBox.Enabled = False
        Dim xlApp As Excel.Application = New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet


        Dim session As Session
        Dim mDBLinkCmpRW As DBLink

        xlWorkBook = xlApp.Workbooks.Open(Path)
        xlWorkSheet = xlWorkBook.Worksheets("Customer Creation Form")

        Try
            session = New Session()
            session.Init("", "XX", "XX1000", "63A") 'first 3 parameter is always like that i dont know why, 4th parameter is Sage Version
            session.Open("ADMIN", "ADMS4G3COM1", DatabaseBox.SelectedItem.ToString(), DateTime.Today, 0) 'Password and Username must be in UPPERCASE
            mDBLinkCmpRW = session.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)

            Dim ARCUSTOMER1header As View
            Dim ARCUSTOMER1detail As View
            Dim ARCUSTSTAT2 As View
            Dim ARCUSTCMT3 As View

            ARCUSTOMER1header = mDBLinkCmpRW.OpenView("AR0024")
            ARCUSTOMER1detail = mDBLinkCmpRW.OpenView("AR0400")
            ARCUSTSTAT2 = mDBLinkCmpRW.OpenView("AR0022")
            ARCUSTCMT3 = mDBLinkCmpRW.OpenView("AR0021")

            ARCUSTOMER1header.Compose({ARCUSTOMER1detail})
            ARCUSTOMER1detail.Compose({ARCUSTOMER1header})

            'INSERT DATA TO ACCPAC

            ARCUSTOMER1header.Init()
            ARCUSTOMER1header.Fields.FieldByName("IDGRP").SetValue("DP", False)
            ARCUSTOMER1header.Fields.FieldByName("IDCUST").SetValue("CUST0001", False)
            ARCUSTOMER1header.Fields.FieldByName("NAMECUST").SetValue(xlWorkSheet.Cells(17, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("TEXTSTRE1").SetValue(xlWorkSheet.Cells(25, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("CODEPSTL").SetValue(xlWorkSheet.Cells(31, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("CODECTRY").SetValue(xlWorkSheet.Cells(31, 5).value, False)
            ARCUSTOMER1header.Fields.FieldByName("TEXTPHON1").SetValue(xlWorkSheet.Cells(33, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("TEXTPHON2").SetValue(xlWorkSheet.Cells(33, 5).value, False)
            ARCUSTOMER1header.Fields.FieldByName("EMAIL2").SetValue(xlWorkSheet.Cells(29, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("CTACPHONE").SetValue(xlWorkSheet.Cells(47, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("EMAIL1").SetValue(xlWorkSheet.Cells(45, 2).value, False)
            ARCUSTOMER1header.Fields.FieldByName("NAMECTAC").SetValue(xlWorkSheet.Cells(41, 2).value, False)
            ARCUSTOMER1header.Insert()

            MsgBox("Customer berhasil ditambah", 0, "Completed")

            SearchButton.Enabled = True
            CancelButton.Enabled = True
            UploadButton.Enabled = True
            DatabaseBox.Enabled = True
            LoadingLabel.Visible = False

            xlWorkBook.Close()
            xlApp.Quit()

        Catch e As NullReferenceException
            MsgBox("Mohon pilih Database", 0, "Error")

            SearchButton.Enabled = True
            CancelButton.Enabled = True
            UploadButton.Enabled = True
            DatabaseBox.Enabled = True
            LoadingLabel.Visible = False

        Catch e As Runtime.InteropServices.COMException
            Dim errors As String = ""

            For k As Integer = 0 To session.Errors.Count() - 1
                errors = errors + session.Errors(k).Message
            Next

            MessageBox.Show(errors)

            End

        Catch e As Exception
            MsgBox("Error" + e.ToString(), 0, "Error")

            End

        End Try
    End Sub


End Class
