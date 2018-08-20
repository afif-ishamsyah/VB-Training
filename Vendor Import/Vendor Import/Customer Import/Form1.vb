Imports System.IO
Imports ACCPAC.Advantage
Imports Microsoft.Office.Interop

Public Class VendorImport

    Dim FileExcel As String
    Dim VendorIDCheck As String
    Dim xlApp As Excel.Application = New Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim session As Session
    Dim mDBLinkCmpRW As DBLink
    Dim csQry As View

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FileNameTextbox.ReadOnly = True
        UploadButton.Enabled = False
        CheckExistButton.Enabled = False
        FirstCharacterTextBox.CharacterCasing = CharacterCasing.Upper
        SearchNameTextBox.CharacterCasing = CharacterCasing.Upper
    End Sub

    Private Sub Connect()
        session = New Session()
        session.Init("", "XX", "XX1000", "63A") 'first 3 parameter is always like that i dont know why, 4th parameter is Sage Version
        session.Open("ADMIN", "ADMS4G3COM1", DatabaseBox.SelectedItem.ToString(), DateTime.Today, 0) 'Password and Username must be in UPPERCASE
        mDBLinkCmpRW = session.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)
        csQry = mDBLinkCmpRW.OpenView("CS0120")
    End Sub

    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
        If SearchDialog.ShowDialog() = DialogResult.OK Then
            FileExcel = SearchDialog.FileName
            FileNameTextbox.Text = Path.GetFileName(FileExcel)
            If FileNameTextbox.Text <> "" Then
                UploadButton.Enabled = True
                CheckExistButton.Enabled = True
            End If
        End If

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
                    xlWorkBook = xlApp.Workbooks.Open(FileExcel)
                    xlWorkSheet = xlWorkBook.Worksheets(1)

                    If xlWorkSheet.Name.ToLower() = "vendor details" Then
                        SendtoSage()
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

    Private Sub CheckExistButton_Click(sender As Object, e As EventArgs) Handles CheckExistButton.Click
        Dim extension As String = Path.GetExtension(FileExcel)

        DisableButton()

        If DatabaseBox.SelectedItem Is Nothing Then
            MsgBox("Mohon pilih Database terlebih dahulu", 0, "Error")
            EnableButton()
        Else
            If File.Exists(FileExcel) Then

                If extension = ".xlsx" Or extension = ".xls" Then
                    xlWorkBook = xlApp.Workbooks.Open(FileExcel)
                    xlWorkSheet = xlWorkBook.Worksheets(1)

                    If xlWorkSheet.Name.ToLower() = "vendor details" Then
                        CheckIdExist()
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


    Private Sub SendtoSage()
        Try
            Connect()

            Dim VENDORID As String = xlWorkSheet.Cells(1, 2).value
            Dim TaxClass As String
            Dim TaxStatus As String

            If xlWorkSheet.Cells(28, 2).value = "PKP" Then
                TaxClass = "2"
                TaxStatus = "1"
            Else
                TaxClass = "1"
                TaxStatus = "0"
            End If

            Dim APVENDOR1header As View
            Dim APVENDOR1detail As View
            Dim APVENDSTAT2 As View
            Dim APVENDCMNT3 As View

            APVENDOR1header = mDBLinkCmpRW.OpenView("AP0015")
            APVENDOR1detail = mDBLinkCmpRW.OpenView("AP0407")
            APVENDSTAT2 = mDBLinkCmpRW.OpenView("AP0019")
            APVENDCMNT3 = mDBLinkCmpRW.OpenView("AP0014")

            APVENDOR1header.Compose({APVENDOR1detail})
            APVENDOR1detail.Compose({APVENDOR1header})

            'Check if customer already exist
            csQry.Browse("select a.VENDORID from APVEN a where a.VENDORID='" + VENDORID + "'", True)
            csQry.InternalSet(256)

            While (csQry.Fetch(False))
                VendorIDCheck = csQry.Fields(0).Value.ToString()
            End While

            If String.Compare(VENDORID, VendorIDCheck) = 0 Then
                MsgBox("Vendor sudah ada sebelumnya", 0, "Error")

                EnableButton()
                xlWorkBook.Close(0)
                xlApp.Quit()
            Else
                'If data not exist, insert customer data to Sage

                APVENDOR1header.Init()
                APVENDOR1header.Fields.FieldByName("IDGRP").SetValue(xlWorkSheet.Cells(23, 2).value, False)
                APVENDOR1header.Fields.FieldByName("VENDORID").SetValue(VENDORID, False)
                APVENDOR1header.Fields.FieldByName("VENDNAME").SetValue(xlWorkSheet.Cells(2, 2).value, False)
                APVENDOR1header.Fields.FieldByName("LEGALNAME").SetValue(xlWorkSheet.Cells(2, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TEXTSTRE1").SetValue(xlWorkSheet.Cells(3, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TEXTSTRE2").SetValue(xlWorkSheet.Cells(4, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TEXTSTRE3").SetValue(xlWorkSheet.Cells(5, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TEXTSTRE4").SetValue(xlWorkSheet.Cells(6, 2).value, False)
                APVENDOR1header.Fields.FieldByName("CODEPSTL").SetValue(xlWorkSheet.Cells(10, 2).value, False)
                APVENDOR1header.Fields.FieldByName("CODECTRY").SetValue(xlWorkSheet.Cells(9, 2).value, False)
                APVENDOR1header.Fields.FieldByName("NAMECITY").SetValue(xlWorkSheet.Cells(7, 2).value, False)
                APVENDOR1header.Fields.FieldByName("CODESTTE").SetValue(xlWorkSheet.Cells(8, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TEXTPHON1").SetValue(xlWorkSheet.Cells(11, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TEXTPHON2").SetValue(xlWorkSheet.Cells(12, 2).value, False)
                APVENDOR1header.Fields.FieldByName("EMAIL2").SetValue(xlWorkSheet.Cells(13, 2).value, False)
                APVENDOR1header.Fields.FieldByName("CTACPHONE").SetValue(xlWorkSheet.Cells(16, 2).value, False)
                APVENDOR1header.Fields.FieldByName("EMAIL1").SetValue(xlWorkSheet.Cells(15, 2).value, False)
                APVENDOR1header.Fields.FieldByName("NAMECTAC").SetValue(xlWorkSheet.Cells(14, 2).value, False)
                APVENDOR1header.Fields.FieldByName("IDTAXREGI1").SetValue(xlWorkSheet.Cells(19, 2).value, False)
                APVENDOR1header.Fields.FieldByName("IDACCTSET").SetValue(xlWorkSheet.Cells(25, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TERMSCODE").SetValue(xlWorkSheet.Cells(24, 2).value, False)
                APVENDOR1header.Fields.FieldByName("SHORTNAME").SetValue(xlWorkSheet.Cells(26, 2).value, False)
                APVENDOR1header.Fields.FieldByName("TAXCLASS1").SetValue(TaxClass, False)

                APVENDOR1detail.Fields.FieldByName("OPTFIELD").SetValue("BKACCT", False)
                APVENDOR1detail.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(21, 2).value, False)
                APVENDOR1detail.Insert()
                APVENDOR1detail.Fields.FieldByName("OPTFIELD").SetValue("BKBENE", False)
                APVENDOR1detail.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(22, 2).value, False)
                APVENDOR1detail.Insert()
                APVENDOR1detail.Fields.FieldByName("OPTFIELD").SetValue("BKNAME", False)
                APVENDOR1detail.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(20, 2).value, False)
                APVENDOR1detail.Insert()
                If DatabaseBox.SelectedItem.ToString().Contains("CMWIDT") Then
                    APVENDOR1detail.Fields.FieldByName("OPTFIELD").SetValue("FINANCENAME", False)
                    APVENDOR1detail.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(27, 2).value, False)
                    APVENDOR1detail.Insert()
                    APVENDOR1detail.Fields.FieldByName("OPTFIELD").SetValue("FINANCEPHONE", False)
                    APVENDOR1detail.Fields.FieldByName("VALIFTEXT").SetValue(xlWorkSheet.Cells(29, 2).value, False)
                    APVENDOR1detail.Insert()
                    APVENDOR1detail.Fields.FieldByName("OPTFIELD").SetValue("PKP", False)
                    APVENDOR1detail.Fields.FieldByName("VALIFBOOL").SetValue(TaxStatus, False)
                    APVENDOR1detail.Insert()
                End If

                APVENDOR1header.Insert()

                MsgBox("Vendor berhasil ditambah", 0, "Completed")

                session.Dispose()
                EnableButton()
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

    Private Sub CheckIdExist()
        Try
            Connect()

            Dim VENDORID As String = xlWorkSheet.Cells(1, 2).value

            'Check if customer already exist
            csQry.Browse("select a.VENDORID from APVEN a where a.VENDORID='" + VENDORID + "'", True)
            csQry.InternalSet(256)

            While (csQry.Fetch(False))
                VendorIDCheck = csQry.Fields(0).Value.ToString()
            End While

            If String.Compare(VENDORID, VendorIDCheck) = 0 Then
                MsgBox("ID Vendor ini sudah digunakan di database. Mohon cek dan ganti ID", 0, "Error")

                EnableButton()
                xlWorkBook.Close(0)
                xlApp.Quit()
            Else
                MsgBox("ID Vendor belum ada di database. Aman untuk digunakan", 0, "Safe")

                EnableButton()
                xlWorkBook.Close(0)
                xlApp.Quit()
            End If

            session.Dispose()

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

    Private Sub SearchIDButton_Click(sender As Object, e As EventArgs) Handles SearchIDButton.Click
        Try
            If DatabaseBox.SelectedItem Is Nothing Then
                MsgBox("Mohon pilih Database terlebih dahulu", 0, "Error")

            ElseIf FirstCharacterTextBox.Text <> "" Then
                DisableButton()
                ResultComboBox.Items.Clear()

                Connect()

                csQry.Browse("select a.VENDORID from APVEN a where a.VENDORID like '" + FirstCharacterTextBox.Text + "%' order by a.VENDORID asc", True)
                csQry.InternalSet(256)

                While (csQry.Fetch(False))
                    ResultComboBox.Items.Add(csQry.Fields(0).Value.ToString())
                End While

                ResultComboBox.SelectedIndex = 0

                session.Dispose()
                EnableButton()
            End If

        Catch ex As Runtime.InteropServices.COMException
            Dim errors As String = ""

            If session.Errors IsNot Nothing Then
                For k As Integer = 0 To session.Errors.Count() - 1
                    errors = errors + session.Errors(k).Message
                Next
            End If

            MsgBox(errors, 0, "Error")
            EnableButton()

        Catch ex As Exception
            MsgBox("Error" + ex.ToString(), 0, "Error")
            EnableButton()
        End Try
    End Sub

    Private Sub SearchNameButton_Click(sender As Object, e As EventArgs) Handles SearchNameButton.Click
        Try
            If DatabaseBox.SelectedItem Is Nothing Then
                MsgBox("Mohon pilih Database terlebih dahulu", 0, "Error")

            ElseIf SearchNameTextBox.Text <> "" Then
                DisableButton()
                VendorNameListView.Clear()
                VendorNameListView.View = Windows.Forms.View.Details
                VendorNameListView.Columns.Add("VendorName", 200, HorizontalAlignment.Center)

                Connect()

                csQry.Browse("select a.VENDNAME from APVEN a where upper(a.VENDNAME) like '%" + SearchNameTextBox.Text + "%'", True)
                csQry.InternalSet(256)

                While (csQry.Fetch(False))
                    Dim row As ListViewItem = New ListViewItem(New String() {csQry.Fields(0).Value.ToString()})
                    VendorNameListView.Items.Add(row)
                End While

                session.Dispose()
                EnableButton()
            End If

        Catch ex As Runtime.InteropServices.COMException
            Dim errors As String = ""

            If session.Errors IsNot Nothing Then
                For k As Integer = 0 To session.Errors.Count() - 1
                    errors = errors + session.Errors(k).Message
                Next
            End If

            MsgBox(errors, 0, "Error")
            EnableButton()

        Catch ex As Exception
            MsgBox("Error" + ex.ToString(), 0, "Error")
            EnableButton()
        End Try
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButtons.Click
        End
    End Sub

    Private Sub EnableButton()
        If FileNameTextbox.Text <> "" Then
            UploadButton.Enabled = True
            CheckExistButton.Enabled = True
        End If
        SearchButton.Enabled = True
        CancelButtons.Enabled = True
        DatabaseBox.Enabled = True
        SearchIDButton.Enabled = True
        SearchNameButton.Enabled = True
    End Sub

    Private Sub DisableButton()
        SearchButton.Enabled = False
        CancelButtons.Enabled = False
        UploadButton.Enabled = False
        DatabaseBox.Enabled = False
        CheckExistButton.Enabled = False
        SearchIDButton.Enabled = False
        SearchNameButton.Enabled = False
    End Sub


End Class
