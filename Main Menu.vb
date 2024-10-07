Public Class Form1
    Dim relevant As Integer
    Dim a As Object
    Dim b As Microsoft.Office.Interop.Excel.Workbook
    Dim c As Microsoft.Office.Interop.Excel.Worksheet
    Dim i As Double
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        MsgBox("Welcome to ultrasonic world. Thanks for your selection. Press OK to continue.", MsgBoxStyle.Information, "welcome")
        a = Microsoft.VisualBasic.Interaction.CreateObject("Excel.Application")
        b = a.workbooks.open(My.Application.Info.DirectoryPath & "\UT Report.xlsx")
        c = b.Worksheets(1)
        lbl_LegReport.Text = ""
        lbl_RelevantReport.Text = ""
        lbl_RegReport.Text = ""
    End Sub

    Private Sub Button1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bot_cal.Click

        Dim pt, pa, rg, wb, pxo, Tj, pcld, renge, s, a, b, c As Double
        Dim leg As String

        If txbx_Tj.Text = "" Or txbx_pcld.Text = "" Or txbx_pt.Text = "" Or txbx_pa.Text = "" Or txbx_rg.Text = "" Or txbx_wb.Text = "" Or txbx_pxo.Text = "" Then
            MsgBox("There is no valid data to do Calculation.", MsgBoxStyle.Information, "warning")
        Else
            pt = Val(txbx_pt.Text)
            pa = Val(txbx_pa.Text) * (System.Math.PI / 180)
            rg = Val(txbx_rg.Text)
            wb = Val(txbx_wb.Text)
            pxo = Val(txbx_pxo.Text)
            Tj = Val(txbx_Tj.Text)
            pcld = Val(txbx_pcld.Text)
            leg = 1

            'Block_v1

            If rbot_v1.Checked = True Then

                a = (2 * pt) / System.Math.Cos(pa)
                b = System.Math.Floor(pt / 100)
                c = b + 1
                renge = pt - (c * 100)
                txbx_r_renge.Text = System.Math.Max(System.Math.Abs(renge) + pt, 200)

                txbx_r_sf.Text = Val(txbx_r_renge.Text) / 10

                txbx_r_fep.Text = System.Math.Round(100 / Val(txbx_r_sf.Text), 1)
                txbx_r_sep.Text = System.Math.Round(200 / Val(txbx_r_sf.Text), 1)

                If txbx_r_sep.Text > 10 Then
                    MsgBox("Seconed echo is out of renge.")
                End If

                txbx_R_hsd.Text = System.Math.Round(pt * System.Math.Tan(pa) + (0.5 * rg), 0)
                txbx_R_fsd.Text = System.Math.Round(2 * pt * System.Math.Tan(pa) + (0.5 * rg) + (0.5 * wb), 0)

                s = Tj * Val(txbx_r_sf.Text)
                txbx_R_d.Text = System.Math.Round(s * System.Math.Cos(pa), 0)
                txbx_R_C.Text = System.Math.Round(s * System.Math.Sin(pa), 0)
                txbx_r_y.Text = System.Math.Round(pcld + pxo - (s * (System.Math.Sin(pa))), 0)

                If (pcld + pxo) <= Val(txbx_R_hsd.Text) Then
                    leg = "1"
                End If
                If (pcld + pxo) >= Val(txbx_R_hsd.Text) And (pcld + pxo) < Val(txbx_R_fsd.Text) Then
                    leg = "2"
                End If
                If (pcld + pxo) >= Val(txbx_R_fsd.Text) Then
                    leg = "Up to 2"
                End If

                lbl_LegReport.Text = "Your probe detect a discontinuity in leg " & leg & "."


            End If


            'Block_v2_r25

            If rbot_v2_r25.Checked = True Then

                a = (2 * pt) / System.Math.Cos(pa)
                b = System.Math.Floor(pt / 25)
                c = b + 1
                renge = pt - (c * 25)
                txbx_r_renge.Text = System.Math.Max(System.Math.Abs(renge) + pt, 100)

                txbx_r_sf.Text = Val(txbx_r_renge.Text) / 10

                txbx_r_fep.Text = System.Math.Round(25 / Val(txbx_r_sf.Text), 1)
                txbx_r_sep.Text = System.Math.Round(100 / Val(txbx_r_sf.Text), 1)

                If txbx_r_sep.Text > 10 Then
                    MsgBox("Seconed echo is out of renge.")
                End If

                txbx_R_hsd.Text = System.Math.Round(pt * System.Math.Tan(pa) + (0.5 * rg), 0)
                txbx_R_fsd.Text = System.Math.Round(2 * pt * System.Math.Tan(pa) + (0.5 * rg) + (0.5 * wb), 0)

                s = Tj * Val(txbx_r_sf.Text)
                txbx_R_d.Text = System.Math.Round(s * System.Math.Cos(pa), 0)
                txbx_R_C.Text = System.Math.Round(s * System.Math.Sin(pa), 0)
                txbx_r_y.Text = System.Math.Round(pcld + pxo - (s * (System.Math.Sin(pa))), 0)

                If (pcld + pxo) <= Val(txbx_R_hsd.Text) Then
                    leg = "1"
                End If
                If (pcld + pxo) >= Val(txbx_R_hsd.Text) And (pcld + pxo) < Val(txbx_R_fsd.Text) Then
                    leg = "2"
                End If
                If (pcld + pxo) >= Val(txbx_R_fsd.Text) Then
                    leg = "Up to 2"
                End If

                lbl_LegReport.Text = "Your probe detect a discontinuity in leg " & leg & "."

            End If

            'Block_V2_r50

            If rbot_v2_r50.Checked = True Then

                a = (2 * pt) / System.Math.Cos(pa)
                b = System.Math.Floor(pt / 50)
                c = b + 1
                renge = pt - (c * 50)
                txbx_r_renge.Text = System.Math.Max(System.Math.Abs(renge) + pt, 125)

                txbx_r_sf.Text = Val(txbx_r_renge.Text) / 10

                txbx_r_fep.Text = System.Math.Round(50 / Val(txbx_r_sf.Text), 1)
                txbx_r_sep.Text = System.Math.Round(125 / Val(txbx_r_sf.Text), 1)

                If txbx_r_sep.Text > 10 Then
                    MsgBox("Seconed echo is out of renge.")
                End If

                txbx_R_hsd.Text = System.Math.Round(pt * System.Math.Tan(pa) + (0.5 * rg), 0)
                txbx_R_fsd.Text = System.Math.Round(2 * pt * System.Math.Tan(pa) + (0.5 * rg) + (0.5 * wb), 0)

                s = Tj * Val(txbx_r_sf.Text)
                txbx_R_d.Text = System.Math.Round(s * System.Math.Cos(pa), 0)
                txbx_R_C.Text = System.Math.Round(s * System.Math.Sin(pa), 0)
                txbx_r_y.Text = System.Math.Round(pcld + pxo - (s * (System.Math.Sin(pa))), 0)

                If (pcld + pxo) <= Val(txbx_R_hsd.Text) Then
                    leg = "1"
                End If
                If (pcld + pxo) >= Val(txbx_R_hsd.Text) And (pcld + pxo) < Val(txbx_R_fsd.Text) Then
                    leg = "2"
                End If
                If (pcld + pxo) >= Val(txbx_R_fsd.Text) Then
                    leg = "Up to 2"
                End If

                lbl_LegReport.Text = "Your probe detect a discontinuity in leg " & leg & "."

            End If

            If System.Math.Abs(Val(txbx_r_y.Text)) < (0.5 * wb) And Val(txbx_R_d.Text) < pt Then
                MsgBox("It is a Relevant Pick.", MsgBoxStyle.Critical, "Information")
                relevant = 1
                lbl_RelevantReport.Text = "It is a Relevant Pick."
            Else
                MsgBox("It is a Non Relevant Pick.", MsgBoxStyle.Critical, "Information")
                relevant = 0
                lbl_RelevantReport.Text = "It is a Non Relevant Pick."
            End If

            txbx_pt.Enabled = False
            txbx_pa.Enabled = False
            txbx_rg.Enabled = False
            txbx_wb.Enabled = False
            txbx_pxo.Enabled = False

            bot_rep.Enabled = True
            bot_sho.Enabled = True
        End If

        

    End Sub

    Private Sub bot_rep_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bot_rep.Click

        If rbot_v1.Checked = True Then
            c.Cells(2, 3) = "V1"
        Else
            c.Cells(2, 3) = "V2"
        End If
        c.Cells(3, 3) = txbx_pt.Text
        c.Cells(4, 3) = txbx_pa.Text
        c.Cells(5, 3) = txbx_rg.Text
        c.Cells(2, 5) = txbx_wb.Text
        c.Cells(3, 5) = txbx_pxo.Text
        c.Cells(4, 5) = txbx_r_renge.Text

        If txbx_Tj.Text = "" Or txbx_pcld.Text = "" Or txbx_R_C.Text = "" Or txbx_r_y.Text = "" Or txbx_R_d.Text = "" Then
            MsgBox("There is no valid data to add.", MsgBoxStyle.Information, "warning")
        Else

            c.Cells((8 + i), 1) = i + 1
            c.Cells((8 + i), 3) = txbx_R_d.Text
            c.Cells((8 + i), 4) = txbx_R_C.Text
            c.Cells((8 + i), 5) = txbx_r_y.Text

            If relevant = 1 Then
                c.Cells((8 + i), 6) = "Relevant"
            Else
                c.Cells((8 + i), 6) = "Non Relevant"
            End If

            i += 1
            lbl_RegReport.Text = i & "of 37 reportes, is/are registration in your list."
            If i > 36 Then
                bot_rep.Enabled = False
                MsgBox("Please print or save the report and then continue.", MsgBoxStyle.Information, "warning")
            End If

        End If
        
    End Sub

    Private Sub bot_sho_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bot_sho.Click
        MsgBox("Please don't close this report until test being finish.", MsgBoxStyle.Information, "Recommend")
        a.visible = True
    End Sub

    Private Sub bot_nrep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bot_nrep.Click

        
        MsgBox("It's maybe take a few minute.", MsgBoxStyle.Information, "notice")

        c.Cells(2, 3) = ""
        c.Cells(3, 3) = ""
        c.Cells(4, 3) = ""
        c.Cells(5, 3) = ""
        c.Cells(2, 5) = ""
        c.Cells(3, 5) = ""
        c.Cells(4, 5) = ""

        c.Cells(8, 1) = ""
        c.Cells(9, 1) = ""
        c.Cells(10, 1) = ""
        c.Cells(11, 1) = ""
        c.Cells(12, 1) = ""
        c.Cells(13, 1) = ""
        c.Cells(14, 1) = ""
        c.Cells(15, 1) = ""
        c.Cells(16, 1) = ""
        c.Cells(17, 1) = ""
        c.Cells(18, 1) = ""
        c.Cells(19, 1) = ""
        c.Cells(20, 1) = ""
        c.Cells(21, 1) = ""
        c.Cells(22, 1) = ""
        c.Cells(23, 1) = ""
        c.Cells(24, 1) = ""
        c.Cells(25, 1) = ""
        c.Cells(26, 1) = ""
        c.Cells(27, 1) = ""
        c.Cells(28, 1) = ""
        c.Cells(29, 1) = ""
        c.Cells(30, 1) = ""
        c.Cells(31, 1) = ""
        c.Cells(32, 1) = ""
        c.Cells(33, 1) = ""
        c.Cells(34, 1) = ""
        c.Cells(35, 1) = ""
        c.Cells(36, 1) = ""
        c.Cells(37, 1) = ""
        c.Cells(38, 1) = ""
        c.Cells(39, 1) = ""
        c.Cells(40, 1) = ""
        c.Cells(41, 1) = ""
        c.Cells(42, 1) = ""
        c.Cells(43, 1) = ""
        c.Cells(44, 1) = ""

        c.Cells(8, 3) = ""
        c.Cells(9, 3) = ""
        c.Cells(10, 3) = ""
        c.Cells(11, 3) = ""
        c.Cells(12, 3) = ""
        c.Cells(13, 3) = ""
        c.Cells(14, 3) = ""
        c.Cells(15, 3) = ""
        c.Cells(16, 3) = ""
        c.Cells(17, 3) = ""
        c.Cells(18, 3) = ""
        c.Cells(19, 3) = ""
        c.Cells(20, 3) = ""
        c.Cells(21, 3) = ""
        c.Cells(22, 3) = ""
        c.Cells(23, 3) = ""
        c.Cells(24, 3) = ""
        c.Cells(25, 3) = ""
        c.Cells(26, 3) = ""
        c.Cells(27, 3) = ""
        c.Cells(28, 3) = ""
        c.Cells(29, 3) = ""
        c.Cells(30, 3) = ""
        c.Cells(31, 3) = ""
        c.Cells(32, 3) = ""
        c.Cells(33, 3) = ""
        c.Cells(34, 3) = ""
        c.Cells(35, 3) = ""
        c.Cells(36, 3) = ""
        c.Cells(37, 3) = ""
        c.Cells(38, 3) = ""
        c.Cells(39, 3) = ""
        c.Cells(40, 3) = ""
        c.Cells(41, 3) = ""
        c.Cells(42, 3) = ""
        c.Cells(43, 3) = ""
        c.Cells(44, 3) = ""

        c.Cells(8, 4) = ""
        c.Cells(9, 4) = ""
        c.Cells(10, 4) = ""
        c.Cells(11, 4) = ""
        c.Cells(12, 4) = ""
        c.Cells(13, 4) = ""
        c.Cells(14, 4) = ""
        c.Cells(15, 4) = ""
        c.Cells(16, 4) = ""
        c.Cells(17, 4) = ""
        c.Cells(18, 4) = ""
        c.Cells(19, 4) = ""
        c.Cells(20, 4) = ""
        c.Cells(21, 4) = ""
        c.Cells(22, 4) = ""
        c.Cells(23, 4) = ""
        c.Cells(24, 4) = ""
        c.Cells(25, 4) = ""
        c.Cells(26, 4) = ""
        c.Cells(27, 4) = ""
        c.Cells(28, 4) = ""
        c.Cells(29, 4) = ""
        c.Cells(30, 4) = ""
        c.Cells(31, 4) = ""
        c.Cells(32, 4) = ""
        c.Cells(33, 4) = ""
        c.Cells(34, 4) = ""
        c.Cells(35, 4) = ""
        c.Cells(36, 4) = ""
        c.Cells(37, 4) = ""
        c.Cells(38, 4) = ""
        c.Cells(39, 4) = ""
        c.Cells(40, 4) = ""
        c.Cells(41, 4) = ""
        c.Cells(42, 4) = ""
        c.Cells(43, 4) = ""
        c.Cells(44, 4) = ""

        c.Cells(8, 5) = ""
        c.Cells(9, 5) = ""
        c.Cells(10, 5) = ""
        c.Cells(11, 5) = ""
        c.Cells(12, 5) = ""
        c.Cells(13, 5) = ""
        c.Cells(14, 5) = ""
        c.Cells(15, 5) = ""
        c.Cells(16, 5) = ""
        c.Cells(17, 5) = ""
        c.Cells(18, 5) = ""
        c.Cells(19, 5) = ""
        c.Cells(20, 5) = ""
        c.Cells(21, 5) = ""
        c.Cells(22, 5) = ""
        c.Cells(23, 5) = ""
        c.Cells(24, 5) = ""
        c.Cells(25, 5) = ""
        c.Cells(26, 5) = ""
        c.Cells(27, 5) = ""
        c.Cells(28, 5) = ""
        c.Cells(29, 5) = ""
        c.Cells(30, 5) = ""
        c.Cells(31, 5) = ""
        c.Cells(32, 5) = ""
        c.Cells(33, 5) = ""
        c.Cells(34, 5) = ""
        c.Cells(35, 5) = ""
        c.Cells(36, 5) = ""
        c.Cells(37, 5) = ""
        c.Cells(38, 5) = ""
        c.Cells(39, 5) = ""
        c.Cells(40, 5) = ""
        c.Cells(41, 5) = ""
        c.Cells(42, 5) = ""
        c.Cells(43, 5) = ""
        c.Cells(44, 5) = ""

        c.Cells(8, 6) = ""
        c.Cells(9, 6) = ""
        c.Cells(10, 6) = ""
        c.Cells(11, 6) = ""
        c.Cells(12, 6) = ""
        c.Cells(13, 6) = ""
        c.Cells(14, 6) = ""
        c.Cells(15, 6) = ""
        c.Cells(16, 6) = ""
        c.Cells(17, 6) = ""
        c.Cells(18, 6) = ""
        c.Cells(19, 6) = ""
        c.Cells(20, 6) = ""
        c.Cells(21, 6) = ""
        c.Cells(22, 6) = ""
        c.Cells(23, 6) = ""
        c.Cells(24, 6) = ""
        c.Cells(25, 6) = ""
        c.Cells(26, 6) = ""
        c.Cells(27, 6) = ""
        c.Cells(28, 6) = ""
        c.Cells(29, 6) = ""
        c.Cells(30, 6) = ""
        c.Cells(31, 6) = ""
        c.Cells(32, 6) = ""
        c.Cells(33, 6) = ""
        c.Cells(34, 6) = ""
        c.Cells(35, 6) = ""
        c.Cells(36, 6) = ""
        c.Cells(37, 6) = ""
        c.Cells(38, 6) = ""
        c.Cells(39, 6) = ""
        c.Cells(40, 6) = ""
        c.Cells(41, 6) = ""
        c.Cells(42, 6) = ""
        c.Cells(43, 6) = ""
        c.Cells(44, 6) = ""
        


        rbot_v1.Enabled = True
        rbot_v2_r50.Enabled = True
        rbot_v2_r25.Enabled = True
        txbx_pt.Enabled = True
        txbx_pa.Enabled = True
        txbx_rg.Enabled = True
        txbx_wb.Enabled = True
        txbx_pxo.Enabled = True

        bot_rep.Enabled = False
        bot_sho.Enabled = False

        i = 1

        MsgBox("Now you can make a new report.", MsgBoxStyle.Information, "notice")

    End Sub

    Private Sub bot_clr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bot_clr.Click
       
        txbx_Tj.Text = ""
        txbx_pcld.Text = ""
        txbx_R_C.Text = ""
        txbx_r_y.Text = ""
        txbx_R_d.Text = ""

    End Sub

    Private Sub bot_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bot_exit.Click
        Me.Visible = False
        Me.Enabled = False
        Global.System.Windows.Forms.Application.Exit()
    End Sub
End Class
