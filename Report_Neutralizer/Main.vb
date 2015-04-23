Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing
Imports Microsoft.Win32

Public Class Main
    Public AppsOffice As String

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TargetKey As RegistryKey
        TargetKey = Registry.ClassesRoot.OpenSubKey("Excel.Application")
        If TargetKey Is Nothing Then
            AppsOffice = "XL_NotInstalled"
        Else
            AppsOffice = "XL_Installed"
            TargetKey.Close()
        End If
    End Sub

    Private Sub btnbrow_outmassrc_Click(sender As Object, e As EventArgs) Handles btnbrow_outmassrc.Click
        Dim outmassrc_path As String
        If OFD_src.ShowDialog() = DialogResult.OK Then
            outmassrc_path = OFD_src.FileName
            txtoutmas_src.Text = outmassrc_path
        End If
        If txtDest_OutMas.Text <> "" Then
            btnNeuOutMas.Enabled = True
        Else
            btnNeuOutMas.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_OutMasDes_Click(sender As Object, e As EventArgs) Handles btnBrow_OutMasDes.Click
        Dim outmasdest_path As String
        Dim datenow As DateTime = DateTime.Now
        SFD_Dest.FileName = SFD_Dest.FileName & datenow.ToString("ddMMyyyy_HHmm")
        If SFD_Dest.ShowDialog() = DialogResult.OK Then
            outmasdest_path = SFD_Dest.FileName
            txtDest_OutMas.Text = outmasdest_path
        End If
        If txtoutmas_src.Text <> "" Then
            btnNeuOutMas.Enabled = True
        Else
            btnNeuOutMas.Enabled = False
        End If
        If txtDest_OutMas.Text <> "" Then
            btnNeuOutMas.Enabled = True
        Else
            btnNeuOutMas.Enabled = False
        End If
    End Sub



    Private Sub btnNeuOutMas_Click(sender As Object, e As EventArgs) Handles btnNeuOutMas.Click
        PicBar.Visible = True
        'prgbar.Visible = True
        BWOutMas.RunWorkerAsync()

    End Sub

    Private Sub BWOutMas_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWOutMas.RunWorkerCompleted
        PicBar.Visible = False
        'prgbar.Visible = False
        btnNeuOutMas.Visible = False
        txtoutmas_src.Text = ""
        txtDest_OutMas.Text = ""
    End Sub

    Private Sub BWOutMas_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWOutMas.DoWork
        Dim objApp As Excel.Application
        Dim objwBook As Excel.Workbook
        Dim objwSheet As Excel.Worksheet
        'System.Threading.Thread.Sleep(100)
        'BWOutMas.ReportProgress(CInt(1 / 100) * 100)

        Try
            objApp = New Excel.Application
            objwBook = objApp.Workbooks.Open(txtoutmas_src.Text)
            objwSheet = objwBook.Worksheets("UID Outlet Master Report")
            'System.Threading.Thread.Sleep(50)
            'BWOutMas.ReportProgress(CInt(5 / 100) * 100)

            objwSheet.UsedRange.UnMerge()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(10 / 100) * 100)
            objwSheet.UsedRange.WrapText = False
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(15 / 100) * 100)
            objwSheet.UsedRange.ColumnWidth = 15
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(20 / 100) * 100)
            objwSheet.UsedRange.RowHeight = 15
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(25 / 100) * 100)

            objwSheet.Range("A1").Value = objwSheet.Range("C1").Value
            objwSheet.Range("A1").EntireRow.Font.Bold = True
            objwSheet.Range("A1").EntireRow.Font.Name = "Calibri"
            objwSheet.Range("A1").EntireRow.Font.Size = "18"
            objwSheet.Range("A1").EntireRow.RowHeight = 20
            Dim conhead1 As String = objwSheet.Range("D4").Value & " " & objwSheet.Range("I4").Value & "; " & objwSheet.Range("L4").Value
            conhead1 = conhead1 & " " & objwSheet.Range("Q4").Value & "; " & objwSheet.Range("U4").Value & " " & objwSheet.Range("Y4").Value
            conhead1 = conhead1 & "; " & objwSheet.Range("AB4").Value & " " & objwSheet.Range("AE4").Value & "; " & objwSheet.Range("AH4").Value
            conhead1 = conhead1 & " " & objwSheet.Range("AL4").Value & "; " & objwSheet.Range("AO4").Value & " " & objwSheet.Range("AS4").Value
            conhead1 = conhead1 & "; " & objwSheet.Range("AV4").Value & " " & objwSheet.Range("AV4").Value
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(30 / 100) * 100)

            Dim conhead2 As String = objwSheet.Range("B7").Value & " " & objwSheet.Range("H7").Value & "; " & objwSheet.Range("L7").Value & " " & objwSheet.Range("O7").Value

            objwSheet.Range("A4").Value = conhead1
            objwSheet.Range("A4").EntireRow.Font.Name = "Calibri"
            objwSheet.Range("A7").Value = conhead2
            objwSheet.Range("A7").EntireRow.Font.Name = "Calibri"
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(35 / 100) * 100)

            Dim rg1 As Excel.Range = objwSheet.Range("B:D")
            rg1.EntireColumn.Select()
            rg1.EntireColumn.Delete()
            'BWOutMas.ReportProgress(CInt(40 / 100) * 100)
            'System.Threading.Thread.Sleep(100)
            Dim rg2 As Excel.Range = objwSheet.Range("C:F")
            rg2.EntireColumn.Select()
            rg2.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(45 / 100) * 100)
            Dim rg3 As Excel.Range = objwSheet.Range("D:E")
            rg3.EntireColumn.Select()
            rg3.EntireColumn.Delete()
            Dim rg4 As Excel.Range = objwSheet.Range("E:H")
            rg4.EntireColumn.Select()
            rg4.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(50 / 100) * 100)
            Dim rg5 As Excel.Range = objwSheet.Range("F:J")
            rg5.EntireColumn.Select()
            rg5.EntireColumn.Delete()
            objwSheet.Range("G:G").EntireColumn.Delete()
            Dim rg6 As Excel.Range = objwSheet.Range("H:J")
            rg6.EntireColumn.Select()
            rg6.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(55 / 100) * 100)
            objwSheet.Range("I:I").EntireColumn.Delete()
            Dim rg7 As Excel.Range = objwSheet.Range("J:K")
            rg7.EntireColumn.Select()
            rg7.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            Dim rg8 As Excel.Range = objwSheet.Range("L:M")
            rg8.EntireColumn.Select()
            rg8.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(60 / 100) * 100)
            Dim rg9 As Excel.Range = objwSheet.Range("M:N")
            rg9.EntireColumn.Select()
            rg9.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(65 / 100) * 100)
            objwSheet.Range("N:N").EntireColumn.Delete()
            objwSheet.Range("O:O").EntireColumn.Delete()
            Dim rg10 As Excel.Range = objwSheet.Range("P:Q")
            rg10.EntireColumn.Select()
            rg10.EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(70 / 100) * 100)
            Dim rg11 As Excel.Range = objwSheet.Range("Q:R")
            rg11.EntireColumn.Select()
            rg11.EntireColumn.Delete()
            System.Threading.Thread.Sleep(100)
            objwSheet.Range("S:S").EntireColumn.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(75 / 100) * 100)
            objwSheet.Range("A2").EntireRow.Delete()
            objwSheet.Range("A4").EntireRow.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(80 / 100) * 100)
            objwSheet.Range("A6").EntireRow.Delete()
            objwSheet.Range("A6").EntireRow.Delete()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(85 / 100) * 100)

            objwBook.SaveAs(txtDest_OutMas.Text)
            objwBook.Close()
            objApp.Quit()
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(90 / 100) * 100)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objwSheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objwBook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objApp)
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(95 / 100) * 100)
            objwSheet = Nothing
            objwBook = Nothing
            objApp = Nothing
            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'System.Threading.Thread.Sleep(100)
            'BWOutMas.ReportProgress(CInt(100 / 100) * 100)

        Catch ex As Exception
            MessageBox.Show("Error on : " & ex.Message.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BWOutMas_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BWOutMas.ProgressChanged
        'Invoke(Sub()
        Me.prgbar.Text = e.ProgressPercentage & "%"
        'End Sub)
    End Sub


    Private Sub btnBrow_LP3_src_Click(sender As Object, e As EventArgs) Handles btnBrow_LP3_src.Click
        Dim lp3srcpath As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            lp3srcpath = OFD_src.FileName
            txtLP3_src.Text = lp3srcpath
        End If
        If txtLP3_dest.Text <> "" Then
            btnNeuLP3.Enabled = True
        Else
            btnNeuLP3.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_LP3_Dest_Click(sender As Object, e As EventArgs) Handles btnBrow_LP3_Dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_LP3_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename
        Dim lp3destpath As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            lp3destpath = SFD_Dest.FileName
            txtLP3_dest.Text = lp3destpath
        End If
        If txtLP3_src.Text <> "" Then
            btnNeuLP3.Enabled = True
        Else
            btnNeuLP3.Enabled = False
        End If
        If txtLP3_dest.Text <> "" Then
            btnNeuLP3.Enabled = True
        Else
            btnNeuLP3.Enabled = False
        End If
    End Sub

    Private Sub btnNeuLP3_Click(sender As Object, e As EventArgs) Handles btnNeuLP3.Click
        If RBAll.Checked = True Then
            GoTo Gothrough
        ElseIf RBRupiah.Checked = True Then
            GoTo Gothrough
        ElseIf RBQty.Checked = True Then
            GoTo Gothrough
        Else
            Exit Sub
        End If
Gothrough: PicBarLP3.Visible = True
        BWLP3.RunWorkerAsync()
    End Sub

    Private Sub BWLP3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWLP3.RunWorkerCompleted
        PicBarLP3.Visible = False
        btnNeuLP3.Enabled = False
        txtLP3_dest.Text = ""
        txtLP3_src.Text = ""
    End Sub

    Private Sub BWLP3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWLP3.DoWork
        Dim xlAppLP3 As Excel.Application
        Dim xlWbookLP3 As Excel.Workbook
        Dim xlWsheetLP3 As Excel.Worksheet

        Select Case LP3RptType
            Case "ALL"
                Try
                    xlAppLP3 = New Excel.Application
                    xlWbookLP3 = xlAppLP3.Workbooks.Open(txtLP3_src.Text)
                    xlWsheetLP3 = xlWbookLP3.Worksheets("UID Weekly Stock and Sales Repo")

                    xlWsheetLP3.UsedRange.UnMerge()
                    xlWsheetLP3.UsedRange.WrapText = False
                    xlWsheetLP3.UsedRange.ColumnWidth = 15
                    xlWsheetLP3.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLP3.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLP3.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLP3.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLP3.Range("A4")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    'xlWsheetLP3.Range("A2").Value = xlWsheetLP3.Range("B2").Value
                    xlWsheetLP3.Range("A2").RowHeight = 27
                    'xlWsheetLP3.Range("A2").Font.Bold = True
                    'xlWsheetLP3.Range("A2").Font.Name = "Calibri"
                    'xlWsheetLP3.Range("A2").Font.Size = "20"

                    Dim paramhead1 As String
                    paramhead1 = xlWsheetLP3.Range("C6").Value & " " & xlWsheetLP3.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLP3.Range("H6").Value & " " & xlWsheetLP3.Range("K6").Value
                    Dim paramhead2 As String
                    paramhead2 = xlWsheetLP3.Range("N6").Value & " " & xlWsheetLP3.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLP3.Range("U6").Value & " " & xlWsheetLP3.Range("Y6").Value

                    xlWsheetLP3.Range("A6").Value = paramhead1
                    xlWsheetLP3.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLP3.Range("A7").Value = paramhead2
                    xlWsheetLP3.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1 As Excel.Range = xlWsheetLP3.Range("B:C")
                    rg1.EntireColumn.Select()
                    rg1.EntireColumn.Delete()
                    Dim rg2 As Excel.Range = xlWsheetLP3.Range("C:F")
                    rg2.EntireColumn.Select()
                    rg2.EntireColumn.Delete()
                    Dim rg3 As Excel.Range = xlWsheetLP3.Range("D:E")
                    rg3.EntireColumn.Select()
                    rg3.EntireColumn.Delete()
                    Dim rg4 As Excel.Range = xlWsheetLP3.Range("E:I")
                    rg4.EntireColumn.Select()
                    rg4.EntireColumn.Delete()
                    Dim rg5 As Excel.Range = xlWsheetLP3.Range("G:I")
                    rg5.EntireColumn.Select()
                    rg5.EntireColumn.Delete()
                    Dim rg6 As Excel.Range = xlWsheetLP3.Range("H:I")
                    rg6.EntireColumn.Select()
                    rg6.EntireColumn.Delete()
                    xlWsheetLP3.Range("I:I").EntireColumn.Delete()

                    Dim lrow1 As Long
                    lrow1 = xlWsheetLP3.Range("X65536").End(Excel.XlDirection.xlUp).Row

                    Dim r As Integer = lrow1 + 3
                    Dim cr As String = "X" & r & ":X65536"

                    Dim cellrange As Excel.Range = xlWsheetLP3.Range(cr)
                    cellrange.Select()
                    cellrange.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    xlWsheetLP3.Range("Y:Y").EntireColumn.Delete()

                    Dim lrow2 As Long
                    lrow2 = xlWsheetLP3.Range("Z1").End(Excel.XlDirection.xlDown).Row

                    Dim r2 As Integer = lrow2 - 3
                    Dim cr2 As String = "Z1:Z" & r2

                    Dim cellrange2 As Excel.Range = xlWsheetLP3.Range(cr2)
                    cellrange2.Select()
                    cellrange2.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim rg7 As Excel.Range = xlWsheetLP3.Range("AA:AB")
                    rg7.EntireColumn.Select()
                    rg7.EntireColumn.Delete()

                    xlWsheetLP3.Range("A3").EntireRow.Delete()

                    Dim lrow3 As Long
                    lrow3 = xlWsheetLP3.Range("B65536").End(Excel.XlDirection.xlUp).Row

                    Dim r3 As Integer = lrow3
                    Dim cr3 As String = "B" & r3 & ":B65536"

                    Dim cellrange3 As Excel.Range = xlWsheetLP3.Range(cr3)
                    cellrange3.Select()
                    cellrange3.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow4, lrow5 As Long
                    lrow4 = xlWsheetLP3.Range("A65536").End(Excel.XlDirection.xlUp).Row

                    Dim r4 As Integer = lrow4 - 1

                    Dim cr4 As String = "A" & r4

                    lrow5 = xlWsheetLP3.Range(cr4).End(Excel.XlDirection.xlUp).Row
                    Dim r5 As Integer = lrow5 - 1
                    Dim r6 As Integer = lrow5 - 2
                    Dim cr5 As String = "A" & r5 & ":A" & r6

                    xlWsheetLP3.Range(cr5).EntireRow.Delete()

                    xlWsheetLP3.Range("AB:AB").EntireColumn.Delete()

                    xlWbookLP3.SaveAs(txtLP3_dest.Text)
                    xlWbookLP3.Close()
                    xlAppLP3.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLP3)

                    xlWsheetLP3 = Nothing
                    xlWbookLP3 = Nothing
                    xlAppLP3 = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "RPH"
                Try
                    xlAppLP3 = New Excel.Application
                    xlWbookLP3 = xlAppLP3.Workbooks.Open(txtLP3_src.Text)
                    xlWsheetLP3 = xlWbookLP3.Worksheets("UID Weekly Stock and Sales Repo")

                    xlWsheetLP3.UsedRange.UnMerge()
                    xlWsheetLP3.UsedRange.WrapText = False
                    xlWsheetLP3.UsedRange.ColumnWidth = 15
                    xlWsheetLP3.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLP3.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLP3.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLP3.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLP3.Range("A4")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetLP3.Range("A2").RowHeight = 27

                    Dim paramhead1 As String
                    paramhead1 = xlWsheetLP3.Range("C6").Value & " " & xlWsheetLP3.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLP3.Range("H6").Value & " " & xlWsheetLP3.Range("K6").Value
                    Dim paramhead2 As String
                    paramhead2 = xlWsheetLP3.Range("N6").Value & " " & xlWsheetLP3.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLP3.Range("U6").Value & " " & xlWsheetLP3.Range("Y6").Value

                    xlWsheetLP3.Range("A6").Value = paramhead1
                    xlWsheetLP3.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLP3.Range("A7").Value = paramhead2
                    xlWsheetLP3.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5 As Excel.Range
                    rg1 = xlWsheetLP3.Range("B:H")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetLP3.Range("C:D")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetLP3.Range("D:H")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetLP3.Range("F:H")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetLP3.Range("G:H")
                    rg5.Select()
                    rg5.Delete()

                    xlWsheetLP3.Range("H:H").EntireColumn.Delete()
                    xlWsheetLP3.Range("A3").EntireRow.Delete()

                    xlWbookLP3.SaveAs(txtLP3_dest.Text)
                    xlWbookLP3.Close()
                    xlAppLP3.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLP3)

                    xlWsheetLP3 = Nothing
                    xlWbookLP3 = Nothing
                    xlAppLP3 = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "QTY"
                Try
                    xlAppLP3 = New Excel.Application
                    xlWbookLP3 = xlAppLP3.Workbooks.Open(txtLP3_src.Text)
                    xlWsheetLP3 = xlWbookLP3.Worksheets("UID Weekly Stock and Sales Repo")

                    xlWsheetLP3.UsedRange.UnMerge()
                    xlWsheetLP3.UsedRange.WrapText = False
                    xlWsheetLP3.UsedRange.ColumnWidth = 15
                    xlWsheetLP3.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetLP3.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetLP3.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetLP3.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetLP3.Range("A4")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetLP3.Range("A2").RowHeight = 27

                    Dim paramhead1 As String
                    paramhead1 = xlWsheetLP3.Range("C6").Value & " " & xlWsheetLP3.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetLP3.Range("H6").Value & " " & xlWsheetLP3.Range("K6").Value
                    Dim paramhead2 As String
                    paramhead2 = xlWsheetLP3.Range("N6").Value & " " & xlWsheetLP3.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetLP3.Range("U6").Value & " " & xlWsheetLP3.Range("Y6").Value

                    xlWsheetLP3.Range("A6").Value = paramhead1
                    xlWsheetLP3.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetLP3.Range("A7").Value = paramhead2
                    xlWsheetLP3.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7 As Excel.Range
                    rg1 = xlWsheetLP3.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetLP3.Range("C:F")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetLP3.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetLP3.Range("E:I")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetLP3.Range("G:I")
                    rg5.Select()
                    rg5.Delete()
                    rg6 = xlWsheetLP3.Range("H:I")
                    rg6.Select()
                    rg6.Delete()

                    xlWsheetLP3.Range("I:I").EntireColumn.Delete()

                    rg7 = xlWsheetLP3.Range("AA:AB")
                    rg7.Select()
                    rg7.Delete()

                    xlWsheetLP3.Range("A3").EntireRow.Delete()

                    xlWbookLP3.SaveAs(txtLP3_dest.Text)
                    xlWbookLP3.Close()
                    xlAppLP3.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookLP3)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppLP3)

                    xlWsheetLP3 = Nothing
                    xlWbookLP3 = Nothing
                    xlAppLP3 = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select


    End Sub

    Dim LP3RptType As String
    Private Sub RBAll_CheckedChanged(sender As Object, e As EventArgs) Handles RBAll.CheckedChanged
        LP3RptType = "ALL"
    End Sub

    Private Sub RBRupiah_CheckedChanged(sender As Object, e As EventArgs) Handles RBRupiah.CheckedChanged
        LP3RptType = "RPH"
    End Sub

    Private Sub RBQty_CheckedChanged(sender As Object, e As EventArgs) Handles RBQty.CheckedChanged
        LP3RptType = "QTY"
    End Sub

    Private Sub btnBrowDistStock_src_Click(sender As Object, e As EventArgs) Handles btnBrowDistStock_src.Click
        Dim diststockpathSrc As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            diststockpathSrc = OFD_src.FileName
            txtDistStock_src.Text = diststockpathSrc
        End If
        If txtDistStock_Dest.Text <> "" Then
            btnNeu_DistStock.Enabled = True
        Else
            btnNeu_DistStock.Enabled = False
        End If
    End Sub

    Private Sub btnBrowDistStock_Dest_Click(sender As Object, e As EventArgs) Handles btnBrowDistStock_Dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DistStock_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename
        Dim DistStockPath_Dest As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            DistStockPath_Dest = SFD_Dest.FileName
            txtDistStock_Dest.Text = DistStockPath_Dest
        End If
        If txtDistStock_src.Text <> "" Then
            btnNeu_DistStock.Enabled = True
        Else
            btnNeu_DistStock.Enabled = False
        End If
        If txtDistStock_Dest.Text <> "" Then
            btnNeu_DistStock.Enabled = True
        Else
            btnNeu_DistStock.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DistStock_Click(sender As Object, e As EventArgs) Handles btnNeu_DistStock.Click
        PicBar_DistStock.Visible = True
        BWDistStock.RunWorkerAsync()
    End Sub

    Private Sub BWDistStock_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDistStock.RunWorkerCompleted
        PicBar_DistStock.Visible = False
        btnNeu_DistStock.Enabled = False
        txtDistStock_Dest.Text = ""
        txtDistStock_src.Text = ""
    End Sub

    Private Sub BWDistStock_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDistStock.DoWork
        Dim xlAppDistStock As Excel.Application
        Dim xlWbookDistStock As Excel.Workbook
        Dim xlWsheetDistStock As Excel.Worksheet

        Try
            xlAppDistStock = New Excel.Application
            xlWbookDistStock = xlAppDistStock.Workbooks.Open(txtDistStock_src.Text)
            xlWsheetDistStock = xlWbookDistStock.Worksheets("UID Distributor Stock Report")

            xlWsheetDistStock.UsedRange.UnMerge()
            xlWsheetDistStock.UsedRange.WrapText = False
            xlWsheetDistStock.UsedRange.ColumnWidth = 15
            xlWsheetDistStock.UsedRange.RowHeight = 15

            Dim rg_head_cut1 As Excel.Range = xlWsheetDistStock.Range("B2")
            Dim rg_head_paste1 As Excel.Range = xlWsheetDistStock.Range("A2")
            rg_head_cut1.Select()
            rg_head_cut1.Cut(rg_head_paste1)

            Dim rg_head_cut2 As Excel.Range = xlWsheetDistStock.Range("B3")
            Dim rg_head_paste2 As Excel.Range = xlWsheetDistStock.Range("A3")
            rg_head_cut2.Select()
            rg_head_cut2.Cut(rg_head_paste2)

            xlWsheetDistStock.Range("A2").RowHeight = 27

            Dim paramhead1, paramhead2, paramhead3 As String
            paramhead1 = xlWsheetDistStock.Range("C5").Value & " " & xlWsheetDistStock.Range("G5").Value & "; "
            paramhead1 = paramhead1 & xlWsheetDistStock.Range("J5").Value & " " & xlWsheetDistStock.Range("M5").Value & "; "
            paramhead1 = paramhead1 & xlWsheetDistStock.Range("V5").Value & " " & xlWsheetDistStock.Range("Z5").Value

            paramhead2 = xlWsheetDistStock.Range("AC5").Value & " " & xlWsheetDistStock.Range("AG5").Value

            paramhead3 = xlWsheetDistStock.Range("C7").Value & " " & xlWsheetDistStock.Range("G7").Value & "; "
            paramhead3 = paramhead3 & xlWsheetDistStock.Range("J7").Value & " " & xlWsheetDistStock.Range("P7").Value & "; "
            paramhead3 = paramhead3 & xlWsheetDistStock.Range("Z7").Value & " " & xlWsheetDistStock.Range("AG7").Value

            xlWsheetDistStock.Range("A5").Value = paramhead1
            xlWsheetDistStock.Range("A5").EntireRow.Font.Name = "Calibri"
            xlWsheetDistStock.Range("A6").Value = paramhead2
            xlWsheetDistStock.Range("A6").EntireRow.Font.Name = "Calibri"
            xlWsheetDistStock.Range("A7").Value = paramhead3
            xlWsheetDistStock.Range("A7").EntireRow.Font.Name = "Calibri"

            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10 As Excel.Range
            rg1 = xlWsheetDistStock.Range("B:C")
            rg1.Select()
            rg1.Delete()
            rg2 = xlWsheetDistStock.Range("C:E")
            rg2.Select()
            rg2.Delete()
            rg3 = xlWsheetDistStock.Range("D:E")
            rg3.Select()
            rg3.Delete()
            rg4 = xlWsheetDistStock.Range("E:F")
            rg4.Select()
            rg4.Delete()
            rg5 = xlWsheetDistStock.Range("F:G")
            rg5.Select()
            rg5.Delete()

            xlWsheetDistStock.Range("G:G").EntireColumn.Delete()
            xlWsheetDistStock.Range("H:H").EntireColumn.Delete()

            rg6 = xlWsheetDistStock.Range("I:K")
            rg6.Select()
            rg6.Delete()

            xlWsheetDistStock.Range("J:J").EntireColumn.Delete()

            rg7 = xlWsheetDistStock.Range("K:L")
            rg7.Select()
            rg7.Delete()
            rg8 = xlWsheetDistStock.Range("L:O")
            rg8.Select()
            rg8.Delete()
            rg9 = xlWsheetDistStock.Range("M:N")
            rg9.Select()
            rg9.Delete()

            rg10 = xlWsheetDistStock.Range("A9:A10")
            rg10.Select()
            xlWsheetDistStock.Range("B9:B10").Copy(rg10)
            xlWsheetDistStock.Range("A9").Value = "Kode Produk"
            xlWsheetDistStock.Range("A9").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetDistStock.Range("A9:A10").Merge()
            xlWsheetDistStock.Range("A9:A10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("B9:B10").Merge()
            xlWsheetDistStock.Range("B9:B10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("C9:C10").Merge()
            xlWsheetDistStock.Range("C9:C10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("D9:D10").Merge()
            xlWsheetDistStock.Range("D9:D10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("E9:G9").Merge()
            xlWsheetDistStock.Range("E9:G9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("H9:H10").Merge()
            xlWsheetDistStock.Range("H9:H10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("H9:H10").WrapText = True
            xlWsheetDistStock.Range("I9:I10").Merge()
            xlWsheetDistStock.Range("I9:I10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("I9:I10").WrapText = True
            xlWsheetDistStock.Range("J9:J10").Merge()
            xlWsheetDistStock.Range("J9:J10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("K9:K10").Merge()
            xlWsheetDistStock.Range("K9:K10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDistStock.Range("L9:L10").Merge()
            xlWsheetDistStock.Range("L9:L10").HorizontalAlignment = Excel.Constants.xlCenter

            xlWbookDistStock.SaveAs(txtDistStock_Dest.Text)
            xlWbookDistStock.Close()
            xlAppDistStock.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDistStock)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDistStock)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDistStock)

            xlWsheetDistStock = Nothing
            xlWbookDistStock = Nothing
            xlAppDistStock = Nothing

            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnBrowProdMas_src_Click(sender As Object, e As EventArgs) Handles btnBrowProdMas_src.Click
        Dim prodmaspathSrc As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            prodmaspathSrc = OFD_src.FileName
            txtProdMas_src.Text = prodmaspathSrc
        End If
        If txtProdMas_dest.Text <> "" Then
            btnNeu_ProdMas.Enabled = True
        Else
            btnNeu_ProdMas.Enabled = False
        End If
    End Sub

    Private Sub btnBrowProdMas_dest_Click(sender As Object, e As EventArgs) Handles btnBrowProdMas_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_ProdMast_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename
        Dim ProdMasPath_Dest As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            ProdMasPath_Dest = SFD_Dest.FileName
            txtProdMas_dest.Text = ProdMasPath_Dest
        End If
        If txtProdMas_src.Text <> "" Then
            btnNeu_ProdMas.Enabled = True
        Else
            btnNeu_ProdMas.Enabled = False
        End If
        If txtProdMas_dest.Text <> "" Then
            btnNeu_ProdMas.Enabled = True
        Else
            btnNeu_ProdMas.Enabled = False
        End If
    End Sub

    Private Sub BWProdMas_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWProdMas.RunWorkerCompleted
        PicBar_ProdMas.Visible = False
        btnNeu_ProdMas.Enabled = False
        txtProdMas_dest.Text = ""
        txtProdMas_src.Text = ""
    End Sub

    Private Sub btnNeu_ProdMas_Click(sender As Object, e As EventArgs) Handles btnNeu_ProdMas.Click
        PicBar_ProdMas.Visible = True
        BWProdMas.RunWorkerAsync()
    End Sub

    Private Sub BWProdMas_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWProdMas.DoWork
        Dim xlAppProdMas As Excel.Application
        Dim xlWbookProdMas As Excel.Workbook
        Dim xlWsheetProdMas As Excel.Worksheet

        Try
            xlAppProdMas = New Excel.Application
            xlWbookProdMas = xlAppProdMas.Workbooks.Open(txtProdMas_src.Text)
            xlWsheetProdMas = xlWbookProdMas.Worksheets("UID Product Master Report")

            xlWsheetProdMas.UsedRange.UnMerge()
            xlWsheetProdMas.UsedRange.WrapText = False
            xlWsheetProdMas.UsedRange.ColumnWidth = 15
            xlWsheetProdMas.UsedRange.RowHeight = 15

            Dim rg_head_cut1 As Excel.Range = xlWsheetProdMas.Range("B2")
            Dim rg_head_paste1 As Excel.Range = xlWsheetProdMas.Range("A2")
            rg_head_cut1.Select()
            rg_head_cut1.Cut(rg_head_paste1)

            Dim rg_head_cut2 As Excel.Range = xlWsheetProdMas.Range("B4")
            Dim rg_head_paste2 As Excel.Range = xlWsheetProdMas.Range("A3")
            rg_head_cut2.Select()
            rg_head_cut2.Cut(rg_head_paste2)

            Dim paramhead1, paramhead2, paramhead3 As String
            paramhead1 = xlWsheetProdMas.Range("C6").Value & " " & xlWsheetProdMas.Range("F6").Value & "; "
            paramhead1 = paramhead1 & xlWsheetProdMas.Range("I6").Value & " " & xlWsheetProdMas.Range("L6").Value & "; "

            paramhead2 = xlWsheetProdMas.Range("R6").Value & " " & xlWsheetProdMas.Range("V6").Value & "; "
            paramhead2 = paramhead2 & xlWsheetProdMas.Range("Y6").Value & " " & xlWsheetProdMas.Range("AB6").Value & "; "

            paramhead3 = xlWsheetProdMas.Range("C8").Value & " " & xlWsheetProdMas.Range("F6").Value & "; "
            paramhead3 = paramhead3 & xlWsheetProdMas.Range("J8").Value & " " & xlWsheetProdMas.Range("L8").Value & "; "
            paramhead3 = paramhead3 & xlWsheetProdMas.Range("R8").Value & " " & xlWsheetProdMas.Range("V8").Value & "; "

            xlWsheetProdMas.Range("A5").Value = paramhead1
            xlWsheetProdMas.Range("A5").EntireRow.Font.Name = "Calibri"
            xlWsheetProdMas.Range("A6").Value = paramhead2
            xlWsheetProdMas.Range("A6").EntireRow.Font.Name = "Calibri"
            xlWsheetProdMas.Range("A7").Value = paramhead3
            xlWsheetProdMas.Range("A7").EntireRow.Font.Name = "Calibri"

            xlWsheetProdMas.Range("A2").RowHeight = 27

            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9 As Excel.Range
            rg1 = xlWsheetProdMas.Range("B:C")
            rg1.Select()
            rg1.Delete()

            rg2 = xlWsheetProdMas.Range("C:J")
            rg2.Select()
            rg2.Delete()

            rg3 = xlWsheetProdMas.Range("D:F")
            rg3.Select()
            rg3.Delete()

            xlWsheetProdMas.Range("E:E").EntireColumn.Delete()

            rg4 = xlWsheetProdMas.Range("F:K")
            rg4.Select()
            rg4.Delete()

            rg5 = xlWsheetProdMas.Range("G:H")
            rg5.Select()
            rg5.Delete()

            xlWsheetProdMas.Range("I:I").EntireColumn.Delete()
            xlWsheetProdMas.Range("AT:AT").EntireColumn.Delete()
            xlWsheetProdMas.Range("A10").EntireRow.Delete()

            rg6 = xlWsheetProdMas.Range("A10:A11")
            xlWsheetProdMas.Range("E10:E11").Copy(rg6)
            xlWsheetProdMas.Range("A10").Value = "SKU"
            xlWsheetProdMas.Range("A10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetProdMas.Range("A10:A11").Merge()
            xlWsheetProdMas.Range("A10:A11").HorizontalAlignment = Excel.Constants.xlCenter

            rg7 = xlWsheetProdMas.Range("B10:B11")
            xlWsheetProdMas.Range("E10:E11").Copy(rg7)
            xlWsheetProdMas.Range("B10").Value = "Description"
            xlWsheetProdMas.Range("B10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetProdMas.Range("B10:B11").Merge()
            xlWsheetProdMas.Range("B10:B11").HorizontalAlignment = Excel.Constants.xlCenter

            rg8 = xlWsheetProdMas.Range("C10:C11")
            xlWsheetProdMas.Range("E10:E11").Copy(rg8)
            xlWsheetProdMas.Range("C10").Value = "Product Type"
            xlWsheetProdMas.Range("C10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetProdMas.Range("C10:C11").Merge()
            xlWsheetProdMas.Range("C10:C11").HorizontalAlignment = Excel.Constants.xlCenter

            rg9 = xlWsheetProdMas.Range("D10:D11")
            xlWsheetProdMas.Range("E10:E11").Copy(rg9)
            xlWsheetProdMas.Range("D10").Value = "PC/CS"
            xlWsheetProdMas.Range("D10").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetProdMas.Range("D10:D11").Merge()
            xlWsheetProdMas.Range("D10:D11").HorizontalAlignment = Excel.Constants.xlCenter

            xlWsheetProdMas.Range("E10:E11").Merge()
            xlWsheetProdMas.Range("E10:E11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("F10:G10").Merge()
            xlWsheetProdMas.Range("F10:G10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("F11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("H10:J10").Merge()
            xlWsheetProdMas.Range("H10:J10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("H11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("K10:N10").Merge()
            xlWsheetProdMas.Range("K10:N10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("O10:P10").Merge()
            xlWsheetProdMas.Range("O10:P10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("R10:S10").Merge()
            xlWsheetProdMas.Range("R10:S10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AO10:AQ10").Merge()
            xlWsheetProdMas.Range("AO10:AQ10").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("T10:T11").Merge()
            xlWsheetProdMas.Range("T10:T11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("U10:U11").Merge()
            xlWsheetProdMas.Range("U10:U11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("V10:V11").Merge()
            xlWsheetProdMas.Range("V10:V11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("W10:W11").Merge()
            xlWsheetProdMas.Range("W10:W11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("V10:V11").Merge()
            xlWsheetProdMas.Range("V10:V11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("X10:X11").Merge()
            xlWsheetProdMas.Range("X10:X11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("Y10:Y11").Merge()
            xlWsheetProdMas.Range("Y10:Y11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("Z10:Z11").Merge()
            xlWsheetProdMas.Range("Z10:Z11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AA10:AA11").Merge()
            xlWsheetProdMas.Range("AA10:AA11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AA10:AA11").WrapText = True
            xlWsheetProdMas.Range("AB10:AB11").Merge()
            xlWsheetProdMas.Range("AB10:AB11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AC10:AC11").Merge()
            xlWsheetProdMas.Range("AC10:AC11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AD10:AD11").Merge()
            xlWsheetProdMas.Range("AD10:AD11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AE10:AE11").Merge()
            xlWsheetProdMas.Range("AE10:AE11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AF10:AF11").Merge()
            xlWsheetProdMas.Range("AF10:AF11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AG10:AG11").Merge()
            xlWsheetProdMas.Range("AG10:AG11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AH10:AH11").Merge()
            xlWsheetProdMas.Range("AH10:AH11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AI10:AI11").Merge()
            xlWsheetProdMas.Range("AI10:AI11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AJ10:AJ11").Merge()
            xlWsheetProdMas.Range("AJ10:AJ11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AK10:AK11").Merge()
            xlWsheetProdMas.Range("AK10:AK11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AL10:Al11").Merge()
            xlWsheetProdMas.Range("AL10:Al11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AM10:AM11").Merge()
            xlWsheetProdMas.Range("AM10:AM11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AN10:AN11").Merge()
            xlWsheetProdMas.Range("AN10:AN11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AR10:AR11").Merge()
            xlWsheetProdMas.Range("AR10:AR11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetProdMas.Range("AS10:AS11").Merge()
            xlWsheetProdMas.Range("AS10:AS11").HorizontalAlignment = Excel.Constants.xlCenter

            xlWbookProdMas.SaveAs(txtProdMas_dest.Text)
            xlWbookProdMas.Close()
            xlAppProdMas.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetProdMas)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookProdMas)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppProdMas)

            xlWsheetProdMas = Nothing
            xlWbookProdMas = Nothing
            xlAppProdMas = Nothing

            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBrow_SPR_src_Click(sender As Object, e As EventArgs) Handles btnBrow_SPR_src.Click
        Dim SPRpathSrc As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            SPRpathSrc = OFD_src.FileName
            txtSPR_src.Text = SPRpathSrc
        End If
        If txtSPR_dest.Text <> "" Then
            btnNeu_SPR.Enabled = True
        Else
            btnNeu_SPR.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_SPR_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_SPR_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_SalesPerf" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename
        Dim SPRpath_Dest As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            SPRpath_Dest = SFD_Dest.FileName
            txtSPR_dest.Text = SPRpath_Dest
        End If
        If txtSPR_src.Text <> "" Then
            btnNeu_SPR.Enabled = True
        Else
            btnNeu_SPR.Enabled = False
        End If
        If txtSPR_dest.Text <> "" Then
            btnNeu_SPR.Enabled = True
        Else
            btnNeu_SPR.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_SPR_Click(sender As Object, e As EventArgs) Handles btnNeu_SPR.Click
        PicBar_SPR.Visible = True
        BWSalesPerf.RunWorkerAsync()
    End Sub

    Private Sub BWSalesPerf_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWSalesPerf.RunWorkerCompleted
        PicBar_SPR.Visible = False
        btnNeu_SPR.Enabled = False
        txtSPR_dest.Text = ""
        txtSPR_src.Text = ""
    End Sub

    Dim SPRreptype As String
    Private Sub RBSPR_Qty_CheckedChanged(sender As Object, e As EventArgs) Handles RBSPR_Qty.CheckedChanged
        SPRreptype = "QTY"
    End Sub

    Private Sub RBSPR_Val_CheckedChanged(sender As Object, e As EventArgs) Handles RBSPR_Val.CheckedChanged
        SPRreptype = "VAL"
    End Sub

    Private Sub RBSPR_Avg_CheckedChanged(sender As Object, e As EventArgs) Handles RBSPR_Avg.CheckedChanged
        SPRreptype = "AVG"
    End Sub

    Private Sub BWSalesPerf_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWSalesPerf.DoWork
        Dim xlAppSPR As Excel.Application
        Dim xlWbookSPR As Excel.Workbook
        Dim xlWsheetSPR As Excel.Worksheet

        Select Case SPRreptype
            Case "QTY"
                'Need Confirmation to Pak Arif for Dynamic Column

            Case "VAL"
                Try
                    xlAppSPR = New Excel.Application
                    xlWbookSPR = xlAppSPR.Workbooks.Open(txtSPR_src.Text)
                    xlWsheetSPR = xlWbookSPR.Worksheets("UID Sales Perfomance Report")

                    xlWsheetSPR.UsedRange.UnMerge()
                    xlWsheetSPR.UsedRange.WrapText = False
                    xlWsheetSPR.UsedRange.ColumnWidth = 15
                    xlWsheetSPR.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetSPR.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetSPR.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetSPR.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetSPR.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetSPR.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2 As String
                    paramhead1 = xlWsheetSPR.Range("C6").Value & " " & xlWsheetSPR.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetSPR.Range("H6").Value & " " & xlWsheetSPR.Range("K6").Value & "; "

                    paramhead2 = xlWsheetSPR.Range("P6").Value & " " & xlWsheetSPR.Range("S6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetSPR.Range("W6").Value & " " & xlWsheetSPR.Range("Z6").Value & "; "

                    xlWsheetSPR.Range("A6").Value = paramhead1
                    xlWsheetSPR.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetSPR.Range("A7").Value = paramhead2
                    xlWsheetSPR.Range("A7").EntireRow.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6 As Excel.Range
                    rg1 = xlWsheetSPR.Range("B:C")
                    rg1.Select()
                    rg1.Delete()

                    rg2 = xlWsheetSPR.Range("C:F")
                    rg2.Select()
                    rg2.Delete()

                    rg3 = xlWsheetSPR.Range("D:E")
                    rg3.Select()
                    rg3.Delete()

                    xlWsheetSPR.Range("E:E").EntireColumn.Delete()

                    rg4 = xlWsheetSPR.Range("F:G")
                    rg4.Select()
                    rg4.Delete()

                    rg5 = xlWsheetSPR.Range("G:H")
                    rg5.Select()
                    rg5.Delete()

                    rg6 = xlWsheetSPR.Range("H:J")
                    rg6.Select()
                    rg6.Delete()

                    xlWsheetSPR.Range("J:J").EntireColumn.Delete()
                    xlWsheetSPR.Range("A4").EntireRow.Delete()

                    xlWbookSPR.SaveAs(txtSPR_dest.Text)
                    xlWbookSPR.Close()
                    xlAppSPR.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetSPR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookSPR)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppSPR)

                    xlWsheetSPR = Nothing
                    xlWbookSPR = Nothing
                    xlAppSPR = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "AVG"
                'Need Confirmation to Pak Arif for Dynamic Column
        End Select

    End Sub

    Private Sub btnBrow_DSM_src_Click(sender As Object, e As EventArgs) Handles btnBrow_DSM_src.Click
        Dim DSMpathSrc As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            DSMpathSrc = OFD_src.FileName
            txtDSM_src.Text = DSMpathSrc
        End If
        If txtDSM_dest.Text <> "" Then
            btnNeu_DSM.Enabled = True
        Else
            btnNeu_DSM.Enabled = False
        End If
    End Sub

    Private Sub btnBrow_DSM_dest_Click(sender As Object, e As EventArgs) Handles btnBrow_DSM_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DailyStockMut_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename
        Dim DSMpath_Dest As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            DSMpath_Dest = SFD_Dest.FileName
            txtDSM_dest.Text = DSMpath_Dest
        End If
        If txtDSM_src.Text <> "" Then
            btnNeu_DSM.Enabled = True
        Else
            btnNeu_DSM.Enabled = False
        End If
        If txtDSM_dest.Text <> "" Then
            btnNeu_DSM.Enabled = True
        Else
            btnNeu_DSM.Enabled = False
        End If
    End Sub

    Private Sub btnNeu_DSM_Click(sender As Object, e As EventArgs) Handles btnNeu_DSM.Click
        PicBar_DSM.Visible = True
        BWDSM.RunWorkerAsync()
    End Sub

    Private Sub BWDSM_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BWDSM.RunWorkerCompleted
        PicBar_DSM.Visible = False
        btnNeu_DSM.Enabled = False
        txtDSM_dest.Text = ""
        txtDSM_src.Text = ""
    End Sub

    Private Sub BWDSM_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWDSM.DoWork
        Dim xlAppDSM As Excel.Application
        Dim xlWbookDSM As Excel.Workbook
        Dim xlWsheetDSM As Excel.Worksheet

        Try
            xlAppDSM = New Excel.Application
            xlWbookDSM = xlAppDSM.Workbooks.Open(txtDSM_src.Text)
            xlWsheetDSM = xlWbookDSM.Worksheets("UID Daily Stock Mutation Report")

            xlWsheetDSM.UsedRange.UnMerge()
            xlWsheetDSM.UsedRange.WrapText = False
            xlWsheetDSM.UsedRange.ColumnWidth = 15
            xlWsheetDSM.UsedRange.RowHeight = 15

            Dim rg_head_cut1 As Excel.Range = xlWsheetDSM.Range("B2")
            Dim rg_head_paste1 As Excel.Range = xlWsheetDSM.Range("A2")
            rg_head_cut1.Select()
            rg_head_cut1.Cut(rg_head_paste1)

            Dim rg_head_cut2 As Excel.Range = xlWsheetDSM.Range("B4")
            Dim rg_head_paste2 As Excel.Range = xlWsheetDSM.Range("A3")
            rg_head_cut2.Select()
            rg_head_cut2.Cut(rg_head_paste2)

            xlWsheetDSM.Range("A2").RowHeight = 27

            Dim paramhead1, paramhead2, paramhead3 As String
            paramhead1 = xlWsheetDSM.Range("C6").Value & " " & xlWsheetDSM.Range("F6").Value & "; "
            paramhead1 = paramhead1 & xlWsheetDSM.Range("J6").Value & " " & xlWsheetDSM.Range("M6").Value

            paramhead2 = xlWsheetDSM.Range("R6").Value & " " & xlWsheetDSM.Range("U6").Value & "; "
            paramhead2 = paramhead2 & xlWsheetDSM.Range("Y6").Value & " " & xlWsheetDSM.Range("AC6").Value

            paramhead3 = xlWsheetDSM.Range("AG6").Value & " " & xlWsheetDSM.Range("AK6").Value & "; "
            paramhead3 = paramhead3 & xlWsheetDSM.Range("AP6").Value & " " & xlWsheetDSM.Range("AS6").Value

            xlWsheetDSM.Range("A5").Value = paramhead1
            xlWsheetDSM.Range("A5").EntireRow.Font.Name = "Calibri"
            xlWsheetDSM.Range("A6").Value = paramhead2
            xlWsheetDSM.Range("A6").EntireRow.Font.Name = "Calibri"
            xlWsheetDSM.Range("A7").Value = paramhead3
            xlWsheetDSM.Range("A7").EntireRow.Font.Name = "Calibri"

            xlWsheetDSM.Range("G8:G9").Copy(xlWsheetDSM.Range("A8:A9"))
            xlWsheetDSM.Range("A8").Value = "PCODE"
            xlWsheetDSM.Range("A8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetDSM.Range("A8:A9").Merge()
            xlWsheetDSM.Range("A8:A9").HorizontalAlignment = Excel.Constants.xlCenter

            Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9, rg10, rg11 As Excel.Range
            rg1 = xlWsheetDSM.Range("B:C")
            rg1.Select()
            rg1.Delete()

            xlWsheetDSM.Range("E8:E9").Copy(xlWsheetDSM.Range("B8:B9"))
            xlWsheetDSM.Range("B8").Value = "Nama Barang"
            xlWsheetDSM.Range("B8").Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White)
            xlWsheetDSM.Range("B8:B9").Merge()
            xlWsheetDSM.Range("B8:B9").HorizontalAlignment = Excel.Constants.xlCenter

            rg2 = xlWsheetDSM.Range("C:D")
            rg2.Select()
            rg2.Delete()

            xlWsheetDSM.Range("C8:C9").Merge()
            xlWsheetDSM.Range("C8:C9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("C8:C9").WrapText = True
            xlWsheetDSM.Range("D8:D9").Merge()
            xlWsheetDSM.Range("D8:D9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("D8:D9").WrapText = True

            rg3 = xlWsheetDSM.Range("E:F")
            rg3.Select()
            rg3.Delete()

            xlWsheetDSM.Range("E8:E9").Merge()
            xlWsheetDSM.Range("E8:E9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("E8:E9").WrapText = True

            rg4 = xlWsheetDSM.Range("F:G")
            rg4.Select()
            rg4.Delete()

            xlWsheetDSM.Range("F8:F9").Merge()
            xlWsheetDSM.Range("F8:F9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("F8:F9").WrapText = True

            xlWsheetDSM.Range("G:G").EntireColumn.Delete()

            rg5 = xlWsheetDSM.Range("H:I")
            rg5.Select()
            rg5.Delete()

            xlWsheetDSM.Range("G8:H8").Merge()
            xlWsheetDSM.Range("G8:H8").HorizontalAlignment = Excel.Constants.xlCenter

            rg6 = xlWsheetDSM.Range("I:J")
            rg6.Select()
            rg6.Delete()

            xlWsheetDSM.Range("I8:I9").Merge()
            xlWsheetDSM.Range("I8:I9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("I8:I9").WrapText = True

            xlWsheetDSM.Range("J8:J9").Merge()
            xlWsheetDSM.Range("J8:J9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("J8:J9").WrapText = True

            rg7 = xlWsheetDSM.Range("K:L")
            rg7.Select()
            rg7.Delete()

            xlWsheetDSM.Range("K8:K9").Merge()
            xlWsheetDSM.Range("K8:K9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("K8:K9").WrapText = True

            rg8 = xlWsheetDSM.Range("L:N")
            rg8.Select()
            rg8.Delete()

            xlWsheetDSM.Range("L8:L9").Merge()
            xlWsheetDSM.Range("L8:L9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("L8:L9").WrapText = True

            xlWsheetDSM.Range("M8:M9").Merge()
            xlWsheetDSM.Range("M8:M9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("M8:M9").WrapText = True

            rg9 = xlWsheetDSM.Range("N:O")
            rg9.Select()
            rg9.Delete()

            xlWsheetDSM.Range("N8:N9").Merge()
            xlWsheetDSM.Range("N8:N9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("N8:N9").WrapText = True

            xlWsheetDSM.Range("O:O").EntireColumn.Delete()

            xlWsheetDSM.Range("O8:O9").Merge()
            xlWsheetDSM.Range("O8:O9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("O8:O9").WrapText = True

            xlWsheetDSM.Range("P:P").EntireColumn.Delete()

            xlWsheetDSM.Range("P8:P9").Merge()
            xlWsheetDSM.Range("P8:P9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("P8:P9").WrapText = True

            xlWsheetDSM.Range("Q8:Q9").Merge()
            xlWsheetDSM.Range("Q8:Q9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("Q8:Q9").WrapText = True

            xlWsheetDSM.Range("R8:R9").Merge()
            xlWsheetDSM.Range("R8:R9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("R8:R9").WrapText = True

            rg10 = xlWsheetDSM.Range("S:T")
            rg10.Select()
            rg10.Delete()

            xlWsheetDSM.Range("S8:S9").Merge()
            xlWsheetDSM.Range("S8:S9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("S8:S9").WrapText = True

            rg11 = xlWsheetDSM.Range("T:U")
            rg11.Select()
            rg11.Delete()

            xlWsheetDSM.Range("T8:T9").Merge()
            xlWsheetDSM.Range("T8:T9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("T8:T9").WrapText = True

            xlWsheetDSM.Range("U8:U9").Merge()
            xlWsheetDSM.Range("U8:U9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("U8:U9").WrapText = True

            xlWsheetDSM.Range("V8:V9").Merge()
            xlWsheetDSM.Range("V8:V9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("V8:V9").WrapText = True

            xlWsheetDSM.Range("W8:W9").Merge()
            xlWsheetDSM.Range("W8:W9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("W8:W9").WrapText = True

            xlWsheetDSM.Range("X:X").EntireColumn.Delete()

            xlWsheetDSM.Range("X8:X9").Merge()
            xlWsheetDSM.Range("X8:X9").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetDSM.Range("X8:X9").WrapText = True

            Dim lrow1 As Long
            lrow1 = xlWsheetDSM.Range("A65536").End(Excel.XlDirection.xlUp).Row

            Dim cr As String = "A" & lrow1 & ":B" & lrow1

            xlWsheetDSM.Range(cr).Merge()
            xlWsheetDSM.Range(cr).HorizontalAlignment = Excel.Constants.xlRight

            xlWbookDSM.SaveAs(txtDSM_dest.Text)
            xlWbookDSM.Close()
            xlAppDSM.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSM)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSM)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSM)

            xlWsheetDSM = Nothing
            xlWbookDSM = Nothing
            xlAppDSM = Nothing

            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Dim DSStype As String
    Private Sub rbDSSall_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSSall.CheckedChanged
        DSStype = "ALL"
    End Sub

    Private Sub rbDSSrupiah_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSSrupiah.CheckedChanged
        DSStype = "RUPIAH"
    End Sub

    Private Sub rbDSSprod_CheckedChanged(sender As Object, e As EventArgs) Handles rbDSSprod.CheckedChanged
        DSStype = "PRODUK"
    End Sub

    Private Sub bntBrowDSS_src_Click(sender As Object, e As EventArgs) Handles bntBrowDSS_src.Click
        Dim DSSpathSrc As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            DSSpathSrc = OFD_src.FileName
            txtDSS_src.Text = DSSpathSrc
        End If
        If txtDSS_dest.Text <> "" Then
            btnNeu_DSS.Enabled = True
        Else
            btnNeu_DSS.Enabled = False
        End If
    End Sub

    Private Sub btnBrowDSS_dest_Click(sender As Object, e As EventArgs) Handles btnBrowDSS_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String = "Neutralize_DailySalesSumm_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename
        Dim DSSpath_Dest As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            DSSpath_Dest = SFD_Dest.FileName
            txtDSS_dest.Text = DSSpath_Dest
        End If
        If txtDSS_src.Text <> "" Then
            btnNeu_DSS.Enabled = True
        Else
            btnNeu_DSS.Enabled = False
        End If
        If txtDSS_dest.Text <> "" Then
            btnNeu_DSS.Enabled = True
        Else
            btnNeu_DSS.Enabled = False
        End If
    End Sub

    Private Sub bwDSS_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwDSS.RunWorkerCompleted
        PicBar_DSS.Visible = False
        btnNeu_DSS.Enabled = False
        txtDSS_dest.Text = ""
        txtDSS_src.Text = ""
    End Sub

    Private Sub bwDSS_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bwDSS.DoWork
        Dim xlAppDSS As Excel.Application
        Dim xlWbookDSS As Excel.Workbook
        Dim xlWsheetDSS As Excel.Worksheet

        Select Case DSStype
            Case "ALL"
                Try
                    xlAppDSS = New Excel.Application
                    xlWbookDSS = xlAppDSS.Workbooks.Open(txtDSS_src.Text)
                    xlWsheetDSS = xlWbookDSS.Worksheets("UID Daily Sales Summary Report")

                    xlWsheetDSS.UsedRange.UnMerge()
                    xlWsheetDSS.UsedRange.WrapText = False
                    xlWsheetDSS.UsedRange.ColumnWidth = 15
                    xlWsheetDSS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetDSS.Range("C6").Value & " " & xlWsheetDSS.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSS.Range("I6").Value & " " & xlWsheetDSS.Range("M6").Value & "; "

                    paramhead2 = xlWsheetDSS.Range("R6").Value & " " & xlWsheetDSS.Range("V6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSS.Range("AE6").Value & " " & xlWsheetDSS.Range("AI6").Value & "; "

                    paramhead3 = xlWsheetDSS.Range("AM6").Value & " " & xlWsheetDSS.Range("AP6").Value

                    xlWsheetDSS.Range("A5").Value = paramhead1
                    xlWsheetDSS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A6").Value = paramhead2
                    xlWsheetDSS.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A7").Value = paramhead3
                    xlWsheetDSS.Range("A7").EntireColumn.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8, rg9 As Excel.Range
                    rg1 = xlWsheetDSS.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetDSS.Range("C:D")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetDSS.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    xlWsheetDSS.Range("E:E").EntireColumn.Delete()
                    xlWsheetDSS.Range("F:F").EntireColumn.Delete()
                    xlWsheetDSS.Range("G:G").EntireColumn.Delete()
                    xlWsheetDSS.Range("I:I").EntireColumn.Delete()
                    rg4 = xlWsheetDSS.Range("K:L")
                    rg4.Select()
                    rg4.Delete()
                    xlWsheetDSS.Range("M:M").EntireColumn.Delete()
                    xlWsheetDSS.Range("N:N").EntireColumn.Delete()
                    xlWsheetDSS.Range("O:O").EntireColumn.Delete()
                    rg5 = xlWsheetDSS.Range("P:Q")
                    rg5.Select()
                    rg5.Delete()
                    rg6 = xlWsheetDSS.Range("Q:R")
                    rg6.Select()
                    rg6.Delete()
                    rg7 = xlWsheetDSS.Range("S:T")
                    rg7.Select()
                    rg7.Delete()
                    rg8 = xlWsheetDSS.Range("T:U")
                    rg8.Select()
                    rg8.Delete()
                    xlWsheetDSS.Range("W:W").EntireColumn.Delete()

                    Dim lrow8 As Long = xlWsheetDSS.Range("A8").End(Excel.XlDirection.xlDown).Row
                    Dim r8a As Integer = lrow8 + 1
                    Dim r8b As Integer = lrow8 + 2
                    Dim cr8 As String = "A" & r8a & ":A" & r8b
                    xlWsheetDSS.Range(cr8).EntireRow.Delete()

                    Dim lrow1 As Long
                    lrow1 = xlWsheetDSS.Range("B65536").End(Excel.XlDirection.xlUp).Row
                    Dim r1 As Integer = lrow1 + 6
                    Dim cr1 As String = "B" & r1 & ":B65536"
                    Dim cellrange1 As Excel.Range = xlWsheetDSS.Range(cr1)
                    cellrange1.Select()
                    cellrange1.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    xlWsheetDSS.Range("C:C").EntireColumn.Delete()
                    xlWsheetDSS.Range("D:D").EntireColumn.Delete()

                    Dim lrow2 As Long
                    lrow2 = xlWsheetDSS.Range("D65536").End(Excel.XlDirection.xlUp).Row
                    Dim r2 As Integer = lrow2 + 6
                    Dim cr2 As String = "D" & r2 & ":D65536"
                    Dim cellrange2 As Excel.Range = xlWsheetDSS.Range(cr2)
                    cellrange2.Select()
                    cellrange2.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    xlWsheetDSS.Range("F:F").EntireColumn.Delete()

                    Dim lrow3 As Long
                    lrow3 = xlWsheetDSS.Range("E65536").End(Excel.XlDirection.xlUp).Row
                    Dim r3 As Integer = lrow3 + 6
                    Dim cr3 As String = "E" & r3 & ":E65536"
                    Dim cellrange3 As Excel.Range = xlWsheetDSS.Range(cr3)
                    cellrange3.Select()
                    cellrange3.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow4 As Long = xlWsheetDSS.Range("E65536").End(Excel.XlDirection.xlUp).Row
                    Dim r4 As Integer = lrow4 + 6
                    Dim cr4 As String = "E" & r4 & ":E65536"
                    Dim cellrange4 As Excel.Range = xlWsheetDSS.Range(cr4)
                    cellrange4.Select()
                    cellrange4.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow9 As Long = xlWsheetDSS.Range("F1").End(Excel.XlDirection.xlDown).Row
                    Dim r9 As Integer = lrow9 + 3
                    Dim cr9 As String = "F" & lrow9 & ":F" & r9

                    Dim lrow10 As Long = xlWsheetDSS.Range("A8").End(Excel.XlDirection.xlDown).Row
                    Dim r10 As Integer = lrow10 - 3
                    Dim cr10 As String = "B" & lrow10 & ":B" & r10

                    xlWsheetDSS.Range(cr9).Copy(xlWsheetDSS.Range(cr10))
                    xlWsheetDSS.Range("F:F").EntireColumn.Delete()

                    Dim lrow5 As Long = xlWsheetDSS.Range("F65536").End(Excel.XlDirection.xlUp).Row
                    Dim r5 As Integer = lrow5 + 6
                    Dim cr5 As String = "E" & r5 & ":F65536"
                    Dim cellrange5 As Excel.Range = xlWsheetDSS.Range(cr5)
                    cellrange5.Select()
                    cellrange5.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow6 As Long = xlWsheetDSS.Range("H1").End(Excel.XlDirection.xlDown).Row
                    Dim r6 As Integer = lrow6 - 1
                    Dim cr6 As String = "H1:H" & r6
                    Dim cellrange6 As Excel.Range = xlWsheetDSS.Range(cr6)
                    cellrange6.Select()
                    cellrange6.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    Dim lrow7 As Long = xlWsheetDSS.Range("J1").End(Excel.XlDirection.xlDown).Row
                    Dim r7 As Integer = lrow7 - 1
                    Dim cr7 As String = "H1:H" & r7
                    Dim cellrange7 As Excel.Range = xlWsheetDSS.Range(cr7)
                    cellrange7.Select()
                    cellrange7.Delete(Shift:=Excel.XlDirection.xlToLeft)

                    rg9 = xlWsheetDSS.Range("Q:T")
                    rg9.Select()
                    rg9.Delete()

                    xlWbookDSS.SaveAs(txtDSS_dest.Text)
                    xlWbookDSS.Close()
                    xlAppDSS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSS)

                    xlWsheetDSS = Nothing
                    xlWbookDSS = Nothing
                    xlAppDSS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                
            Case "RUPIAH"
                Try
                    xlAppDSS = New Excel.Application
                    xlWbookDSS = xlAppDSS.Workbooks.Open(txtDSS_src.Text)
                    xlWsheetDSS = xlWbookDSS.Worksheets("UID Daily Sales Summary Report")

                    xlWsheetDSS.UsedRange.UnMerge()
                    xlWsheetDSS.UsedRange.WrapText = False
                    xlWsheetDSS.UsedRange.ColumnWidth = 15
                    xlWsheetDSS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetDSS.Range("C6").Value & " " & xlWsheetDSS.Range("E6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSS.Range("H6").Value & " " & xlWsheetDSS.Range("K6").Value & "; "

                    paramhead2 = xlWsheetDSS.Range("N6").Value & " " & xlWsheetDSS.Range("P6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSS.Range("U6").Value & " " & xlWsheetDSS.Range("X6").Value & "; "

                    paramhead3 = xlWsheetDSS.Range("AB6").Value & " " & xlWsheetDSS.Range("AE6").Value

                    xlWsheetDSS.Range("A5").Value = paramhead1
                    xlWsheetDSS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A6").Value = paramhead2
                    xlWsheetDSS.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A7").Value = paramhead3
                    xlWsheetDSS.Range("A7").EntireColumn.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6, rg7, rg8 As Excel.Range
                    rg1 = xlWsheetDSS.Range("B:E")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetDSS.Range("C:E")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetDSS.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    rg4 = xlWsheetDSS.Range("E:G")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetDSS.Range("F:G")
                    rg5.Select()
                    rg5.Delete()
                    xlWsheetDSS.Range("G:G").EntireColumn.Delete()
                    rg6 = xlWsheetDSS.Range("H:I")
                    rg6.Select()
                    rg6.Delete()
                    rg7 = xlWsheetDSS.Range("J:K")
                    rg7.Select()
                    rg7.Delete()
                    rg8 = xlWsheetDSS.Range("K:L")
                    rg8.Select()
                    rg8.Delete()
                    xlWsheetDSS.Range("N:N").EntireColumn.Delete()
                    xlWsheetDSS.Range("Q:Q").EntireColumn.Delete()

                    xlWbookDSS.SaveAs(txtDSS_dest.Text)
                    xlWbookDSS.Close()
                    xlAppDSS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSS)

                    xlWsheetDSS = Nothing
                    xlWbookDSS = Nothing
                    xlAppDSS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Case "PRODUK"
                Try
                    xlAppDSS = New Excel.Application
                    xlWbookDSS = xlAppDSS.Workbooks.Open(txtDSS_src.Text)
                    xlWsheetDSS = xlWbookDSS.Worksheets("UID Daily Sales Summary Report")

                    xlWsheetDSS.UsedRange.UnMerge()
                    xlWsheetDSS.UsedRange.WrapText = False
                    xlWsheetDSS.UsedRange.ColumnWidth = 15
                    xlWsheetDSS.UsedRange.RowHeight = 15

                    Dim rg_head_cut1 As Excel.Range = xlWsheetDSS.Range("B2")
                    Dim rg_head_paste1 As Excel.Range = xlWsheetDSS.Range("A2")
                    rg_head_cut1.Select()
                    rg_head_cut1.Cut(rg_head_paste1)

                    Dim rg_head_cut2 As Excel.Range = xlWsheetDSS.Range("B4")
                    Dim rg_head_paste2 As Excel.Range = xlWsheetDSS.Range("A3")
                    rg_head_cut2.Select()
                    rg_head_cut2.Cut(rg_head_paste2)

                    xlWsheetDSS.Range("A2").RowHeight = 27

                    Dim paramhead1, paramhead2, paramhead3 As String
                    paramhead1 = xlWsheetDSS.Range("C6").Value & " " & xlWsheetDSS.Range("F6").Value & "; "
                    paramhead1 = paramhead1 & xlWsheetDSS.Range("H6").Value & " " & xlWsheetDSS.Range("K6").Value & "; "

                    paramhead2 = xlWsheetDSS.Range("O6").Value & " " & xlWsheetDSS.Range("S6").Value & "; "
                    paramhead2 = paramhead2 & xlWsheetDSS.Range("Z6").Value & " " & xlWsheetDSS.Range("AC6").Value & "; "

                    paramhead3 = xlWsheetDSS.Range("AE6").Value & " " & xlWsheetDSS.Range("AG6").Value

                    xlWsheetDSS.Range("A5").Value = paramhead1
                    xlWsheetDSS.Range("A5").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A6").Value = paramhead2
                    xlWsheetDSS.Range("A6").EntireRow.Font.Name = "Calibri"
                    xlWsheetDSS.Range("A7").Value = paramhead3
                    xlWsheetDSS.Range("A7").EntireColumn.Font.Name = "Calibri"

                    Dim rg1, rg2, rg3, rg4, rg5, rg6 As Excel.Range
                    rg1 = xlWsheetDSS.Range("B:C")
                    rg1.Select()
                    rg1.Delete()
                    rg2 = xlWsheetDSS.Range("C:F")
                    rg2.Select()
                    rg2.Delete()
                    rg3 = xlWsheetDSS.Range("D:E")
                    rg3.Select()
                    rg3.Delete()
                    xlWsheetDSS.Range("E:E").EntireColumn.Delete()

                    Dim lrow1 As Long = xlWsheetDSS.Range("G1").End(Excel.XlDirection.xlDown).Row
                    Dim r1 As Integer = lrow1 + 3
                    Dim cr1 As String = "G" & lrow1 & ":G" & r1
                    Dim cr2 As String = "B" & lrow1 & ":B" & r1

                    Dim cellrange As Excel.Range = xlWsheetDSS.Range(cr1)
                    cellrange.Select()
                    cellrange.Copy(xlWsheetDSS.Range(cr2))

                    rg4 = xlWsheetDSS.Range("F:G")
                    rg4.Select()
                    rg4.Delete()
                    rg5 = xlWsheetDSS.Range("G:H")
                    rg5.Select()
                    rg5.Delete()
                    xlWsheetDSS.Range("H:H").EntireColumn.Delete()
                    xlWsheetDSS.Range("I:I").EntireColumn.Delete()
                    rg6 = xlWsheetDSS.Range("J:S")
                    rg6.Select()
                    rg6.Delete()

                    xlWbookDSS.SaveAs(txtDSS_dest.Text)
                    xlWbookDSS.Close()
                    xlAppDSS.Quit()

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookDSS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppDSS)

                    xlWsheetDSS = Nothing
                    xlWbookDSS = Nothing
                    xlAppDSS = Nothing

                    MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

    Private Sub btnNeu_DSS_Click(sender As Object, e As EventArgs) Handles btnNeu_DSS.Click
        PicBar_DSS.Visible = True
        bwDSS.RunWorkerAsync()
    End Sub

    Private Sub btnBrowProm_src_Click(sender As Object, e As EventArgs) Handles btnBrowProm_src.Click
        Dim PrmpathSrc As String
        If OFD_src.ShowDialog = DialogResult.OK Then
            PrmpathSrc = OFD_src.FileName
            txtPromo_src.Text = PrmpathSrc
        End If
        If txtProm_dest.Text <> "" Then
            btnNeu_Promo.Enabled = True
        Else
            btnNeu_Promo.Enabled = False
        End If
    End Sub

    Private Sub btnBrowPromo_dest_Click(sender As Object, e As EventArgs) Handles btnBrowPromo_dest.Click
        Dim datenow As DateTime = DateTime.Now
        Dim filename As String
        filename = "Neutralize_ListOfPromo_" & datenow.ToString("ddMMyyyy_HHmm")
        SFD_Dest.FileName = filename

        Dim Prompath_Dest As String
        If SFD_Dest.ShowDialog = DialogResult.OK Then
            Prompath_Dest = SFD_Dest.FileName
            txtProm_dest.Text = Prompath_Dest
        End If
        If txtPromo_src.Text <> "" Then
            btnNeu_Promo.Enabled = True
        Else
            btnNeu_Promo.Enabled = False
        End If
        If txtProm_dest.Text <> "" Then
            btnNeu_Promo.Enabled = True
        Else
            btnNeu_Promo.Enabled = False
        End If
    End Sub

    Private Sub bwPromo_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bwPromo.DoWork
        Dim xlAppPRM As Excel.Application
        Dim xlWbookPRM As Excel.Workbook
        Dim xlWsheetPRM As Excel.Worksheet

        Try
            xlAppPRM = New Excel.Application
            xlWbookPRM = xlAppPRM.Workbooks.Open(txtPromo_src.Text)
            xlWsheetPRM = xlWbookPRM.Worksheets("UID List Of Promotion Report")

            xlWsheetPRM.UsedRange.UnMerge()
            xlWsheetPRM.UsedRange.WrapText = False
            xlWsheetPRM.UsedRange.ColumnWidth = 15
            xlWsheetPRM.UsedRange.RowHeight = 15

            Dim rg_head_cut1 As Excel.Range = xlWsheetPRM.Range("B2")
            Dim rg_head_paste1 As Excel.Range = xlWsheetPRM.Range("A2")
            rg_head_cut1.Select()
            rg_head_cut1.Cut(rg_head_paste1)

            Dim rg_head_cut2 As Excel.Range = xlWsheetPRM.Range("B4")
            Dim rg_head_paste2 As Excel.Range = xlWsheetPRM.Range("A3")
            rg_head_cut2.Select()
            rg_head_cut2.Cut(rg_head_paste2)

            xlWsheetPRM.Range("A2").RowHeight = 27

            Dim paramhead1, paramhead2, paramhead3 As String
            paramhead1 = xlWsheetPRM.Range("C6").Value & " " & xlWsheetPRM.Range("F6").Value & "; "
            paramhead1 = paramhead1 & xlWsheetPRM.Range("H6").Value & " " & xlWsheetPRM.Range("J6").Value & "; "

            paramhead2 = xlWsheetPRM.Range("O6").Value & " " & xlWsheetPRM.Range("R6").Value & "; "
            paramhead2 = paramhead2 & xlWsheetPRM.Range("U6").Value & " " & xlWsheetPRM.Range("X6").Value & "; "

            paramhead3 = xlWsheetPRM.Range("C8").Value & " " & xlWsheetPRM.Range("F8").Value

            xlWsheetPRM.Range("A5").Value = paramhead1
            xlWsheetPRM.Range("A5").EntireRow.Font.Name = "Calibri"
            xlWsheetPRM.Range("A6").Value = paramhead2
            xlWsheetPRM.Range("A6").EntireRow.Font.Name = "Calibri"
            xlWsheetPRM.Range("A7").Value = paramhead3
            xlWsheetPRM.Range("A7").EntireRow.Font.Name = "Calibri"

            Dim rg1, rg2, rg3, rg4, rg5, rg6 As Excel.Range
            rg1 = xlWsheetPRM.Range("B:C")
            rg1.Select()
            rg1.Delete()
            rg2 = xlWsheetPRM.Range("C:H")
            rg2.Select()
            rg2.Delete()
            rg3 = xlWsheetPRM.Range("D:E")
            rg3.Select()
            rg3.Delete()
            rg4 = xlWsheetPRM.Range("E:F")
            rg4.Select()
            rg4.Delete()
            xlWsheetPRM.Range("F:F").EntireColumn.Delete()
            rg5 = xlWsheetPRM.Range("G:H")
            rg5.Select()
            rg5.Delete()
            rg6 = xlWsheetPRM.Range("H:J")
            rg6.Select()
            rg6.Delete()
            xlWsheetPRM.Range("J:J").EntireColumn.Delete()

            xlWsheetPRM.Range("A11:A12").Merge()
            xlWsheetPRM.Range("A11:A12").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("B11:B12").Merge()
            xlWsheetPRM.Range("B11:B12").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("C11:C12").Merge()
            xlWsheetPRM.Range("C11:C12").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("D11:D12").Merge()
            xlWsheetPRM.Range("D11:D12").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("E11:F11").Merge()
            xlWsheetPRM.Range("E11:F11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("G11:I11").Merge()
            xlWsheetPRM.Range("G11:I11").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("J11:J12").Merge()
            xlWsheetPRM.Range("J11:J12").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("K11:K12").Merge()
            xlWsheetPRM.Range("K11:K12").HorizontalAlignment = Excel.Constants.xlCenter
            xlWsheetPRM.Range("L11:M11").Merge()
            xlWsheetPRM.Range("L11:M11").HorizontalAlignment = Excel.Constants.xlCenter

            xlWbookPRM.SaveAs(txtProm_dest.Text)
            xlWbookPRM.Close()
            xlAppPRM.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWsheetPRM)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWbookPRM)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlAppPRM)

            xlWsheetPRM = Nothing
            xlWbookPRM = Nothing
            xlAppPRM = Nothing

            MessageBox.Show("Neutralize Completed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message.ToString, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub bwPromo_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwPromo.RunWorkerCompleted
        PicBar_Promo.Visible = False
        btnNeu_Promo.Enabled = False
        txtProm_dest.Text = ""
        txtPromo_src.Text = ""
    End Sub

    Private Sub btnNeu_Promo_Click(sender As Object, e As EventArgs) Handles btnNeu_Promo.Click
        PicBar_Promo.Visible = True
        bwPromo.RunWorkerAsync()
    End Sub
End Class
