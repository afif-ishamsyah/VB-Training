Imports System.IO
Imports ACCPAC.Advantage
Imports Microsoft.Office.Interop

Public Class CustomerImport

    Dim FileExcel As String
    Dim IDCustomerList As New List(Of String)
    Dim xlApp As Excel.Application = New Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FileNameTextbox.ReadOnly = True
        UploadButton.Enabled = False
        LoadingLabel.Visible = False
    End Sub

    Private Sub UploadButton_Click(sender As Object, e As EventArgs) Handles UploadButton.Click
        Dim extension As String = Path.GetExtension(FileExcel)

        DisableButton()

        If DatabaseBox.SelectedItem Is Nothing Then
            MsgBox("Mohon pilih Database terlebih dahulu", 0, "Error")
            EnableButton()
        Else
            If File.Exists(FileExcel) Then

                If extension = ".xlsx" Or extension = ".xls" Then
                    LoadingLabel.Visible = True
                    xlWorkBook = xlApp.Workbooks.Open(FileExcel)
                    xlWorkSheet = xlWorkBook.Worksheets(1)

                    If xlWorkSheet.Name.ToLower() = "customer creation form" Then
                        SenDatoSage()
                    Else
                        MsgBox("File yang dipilih salah", 0, "Error")

                        xlWorkBook.Close(0)
                        xlApp.Quit()
                        EnableButton()
                    End If

                Else
                    MsgBox("File harus bereskstensi Microsoft Excel (.xlsx or .xls)", 0, "Wrong File Type")
                    EnableButton()
                End If

            Else
                MsgBox("File tidak ditemukan. Silahkan lakukan kembali pencarian File", 0, "Error")
                EnableButton()
            End If

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

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelsButton.Click
        End
    End Sub


    Private Sub SenDatoSage()
        Dim session As Session = New Session()
        Dim mDBLinkCmpRW As DBLink
        Try
            session.Init("", "XX", "XX1000", "63A") 'first 3 parameter is always like that i dont know why, 4th parameter is Sage Version
            session.Open("ADMIN", "ADMS4G3COM1", DatabaseBox.SelectedItem.ToString(), DateTime.Today, 0) 'Password and Username must be in UPPERCASE
            mDBLinkCmpRW = session.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)
            Dim IDCUST As String = xlWorkSheet.Cells(23, 3).value

            'CUSTOMER------------------------------------------------------------------
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

            'Check if customer already exist
            Dim searchFilter As String = "IDCUST LIKE %" + IDCUST + "%"

            ARCUSTOMER1header.FilterSelect(searchFilter, True, 0, ViewFilterOrigin.FromStart)
            While ARCUSTOMER1header.FilterFetch(False)
                IDCustomerList.Add(ARCUSTOMER1header.Fields.FieldByName("IDCUST").Value)
            End While

            If IDCustomerList.Contains(IDCUST) Then
                MsgBox("Customer sudah ada sebelumnya", 0, "Error")

                EnableButton()
                LoadingLabel.Visible = False

                xlWorkBook.Close(0)
                xlApp.Quit()
            Else
                'If data not exist, insert customer data to Sage

                ARCUSTOMER1header.Init()
                ARCUSTOMER1header.Fields.FieldByName("IDGRP").SetValue(xlWorkSheet.Cells(23, 8).value, False)
                ARCUSTOMER1header.Fields.FieldByName("IDCUST").SetValue(IDCUST, False)
                ARCUSTOMER1header.Fields.FieldByName("NAMECUST").SetValue(xlWorkSheet.Cells(27, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("TEXTSTRE1").SetValue(xlWorkSheet.Cells(31, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("TEXTSTRE2").SetValue(xlWorkSheet.Cells(33, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("TEXTSTRE3").SetValue(xlWorkSheet.Cells(35, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("TEXTSTRE4").SetValue(xlWorkSheet.Cells(37, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("CODEPSTL").SetValue(xlWorkSheet.Cells(45, 9).value, False)
                ARCUSTOMER1header.Fields.FieldByName("CODECTRY").SetValue(xlWorkSheet.Cells(45, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("NAMECITY").SetValue(xlWorkSheet.Cells(43, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("CODESTTE").SetValue(xlWorkSheet.Cells(43, 9).value, False)
                ARCUSTOMER1header.Fields.FieldByName("TEXTPHON1").SetValue(xlWorkSheet.Cells(39, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("TEXTPHON2").SetValue(xlWorkSheet.Cells(39, 9).value, False)
                ARCUSTOMER1header.Fields.FieldByName("EMAIL2").SetValue(xlWorkSheet.Cells(41, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("CTACPHONE").SetValue(xlWorkSheet.Cells(75, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("EMAIL1").SetValue(xlWorkSheet.Cells(77, 9).value, False)
                ARCUSTOMER1header.Fields.FieldByName("NAMECTAC").SetValue(xlWorkSheet.Cells(71, 3).value, False)
                ARCUSTOMER1header.Fields.FieldByName("IDTAXREGI1").SetValue(xlWorkSheet.Cells(54, 3).value, False)

                ARCUSTOMER1detail.Fields.FieldByName("OPTFIELD").SetValue("CUSTNAME2", False)
                ARCUSTOMER1detail.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(29, 3).value, False)
                ARCUSTOMER1detail.Insert()

                ARCUSTOMER1header.Insert()

                'SHIPTOLOACTION------------------------------------------------------------------

                Dim ARCUSTSHIP2header As View
                Dim ARCUSTSHIP2detailFields As View

                ARCUSTSHIP2header = mDBLinkCmpRW.OpenView("AR0023")
                ARCUSTSHIP2detailFields = mDBLinkCmpRW.OpenView("AR0412")

                ARCUSTSHIP2header.Compose({ARCUSTSHIP2detailFields})
                ARCUSTSHIP2detailFields.Compose({ARCUSTSHIP2header})

                ARCUSTSHIP2header.Init()
                ARCUSTSHIP2header.Fields.FieldByName("IDCUST").SetValue(IDCUST, False)
                ARCUSTSHIP2header.Fields.FieldByName("IDCUSTSHPT").SetValue("NPWP", False)
                ARCUSTSHIP2header.Fields.FieldByName("NAMELOCN").SetValue(xlWorkSheet.Cells(27, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("TEXTSTRE1").SetValue(xlWorkSheet.Cells(56, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("TEXTSTRE2").SetValue(xlWorkSheet.Cells(58, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("TEXTSTRE3").SetValue(xlWorkSheet.Cells(60, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("TEXTSTRE4").SetValue(xlWorkSheet.Cells(62, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("CODECTRY").SetValue(xlWorkSheet.Cells(45, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("CODESTTE").SetValue(xlWorkSheet.Cells(64, 9).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("NAMECITY").SetValue(xlWorkSheet.Cells(64, 3).value, False)
                ARCUSTSHIP2header.Fields.FieldByName("CODETERR").SetValue("0" & xlWorkSheet.Cells(66, 3).value, False)

                ARCUSTSHIP2detailFields.Fields.FieldByName("OPTFIELD").SetValue("CUSTNAME2", False)
                ARCUSTSHIP2detailFields.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(29, 3).value, False)
                ARCUSTSHIP2detailFields.Insert()

                ARCUSTSHIP2header.Insert()

                MsgBox("Customer berhasil ditambah", 0, "Completed")

                EnableButton()
                LoadingLabel.Visible = False

                xlWorkBook.Close(0)
                xlApp.Quit()
            End If

        Catch e As Runtime.InteropServices.COMException
            Dim errors As String = ""

            If session.Errors IsNot Nothing Then
                For k As Integer = 0 To session.Errors.Count() - 1
                    errors = errors + session.Errors(k).Message
                Next
            Else
                errors = errors + "File yang dipilih salah"
            End If

            MsgBox(errors, 0, "Error")

            xlWorkBook.Close(0)
            xlApp.Quit()
            EnableButton()

        Catch e As Exception
            MsgBox("Error" + e.ToString(), 0, "Error")

            xlWorkBook.Close(0)
            xlApp.Quit()
            EnableButton()

        End Try
    End Sub

    Private Sub EnableButton()
        SearchButton.Enabled = True
        CancelsButton.Enabled = True
        UploadButton.Enabled = True
        DatabaseBox.Enabled = True
    End Sub

    Private Sub DisableButton()
        SearchButton.Enabled = False
        CancelsButton.Enabled = False
        UploadButton.Enabled = False
        DatabaseBox.Enabled = False
    End Sub

End Class
