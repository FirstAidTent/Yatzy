Public Class Form1
    Dim slag As Integer
    Dim antal As Integer
    Dim t1, t2, t3, t4, t5 As Integer
    Dim s1, s2, s3, s4, s5 As Boolean
    Dim summa, totalsumma As Integer

    Private Sub Rensa()
        lbl_summa.Text = summa
        lbl_total.Text = totalsumma
        antal = 0
        cmd_slå.Enabled = True
        slag = 0
        lbl_första.Text = ""
        lbl_andra.Text = ""
        lbl_tredje.Text = ""
        lbl_fjärde.Text = ""
        lbl_femte.Text = ""
        t1 = 0
        t2 = 0
        t3 = 0
        t4 = 0
        t5 = 0
        s1 = False
        s2 = False
        s3 = False
        s4 = False
        s5 = False
        cmd_spara1.Text = "Spara"
        cmd_spara2.Text = "Spara"
        cmd_spara3.Text = "Spara"
        cmd_spara4.Text = "Spara"
        cmd_spara5.Text = "Spara"
        If lbl_ett.Text <> "" And lbl_två.Text <> "" And lbl_tre.Text <> "" And lbl_fyra.Text <> "" And lbl_fem.Text <> "" And lbl_sex.Text <> "" And lbl_bonus.Text = "" Then
            If summa >= 63 Then
                lbl_bonus.Text = "50"
                totalsumma = totalsumma + 50
                lbl_total.Text = totalsumma
            Else
                lbl_bonus.Text = "0"
            End If
        End If
    End Sub

    Private Sub Rensa2()
        antal = 0
        cmd_slå.Enabled = True
        slag = 0
        lbl_första.Text = ""
        lbl_andra.Text = ""
        lbl_tredje.Text = ""
        lbl_fjärde.Text = ""
        lbl_femte.Text = ""
        t1 = 0
        t2 = 0
        t3 = 0
        t4 = 0
        t5 = 0
        s1 = False
        s2 = False
        s3 = False
        s4 = False
        s5 = False
        cmd_spara1.Text = "Spara"
        cmd_spara2.Text = "Spara"
        cmd_spara3.Text = "Spara"
        cmd_spara4.Text = "Spara"
        cmd_spara5.Text = "Spara"
    End Sub

    Private Sub cmd_slå_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_slå.Click
        Randomize()
        slag = slag + 1
        If Not s1 Then
            t1 = Int(1 + 6 * Rnd())
            lbl_första.Text = t1
        End If
        If Not s2 Then
            t2 = Int(1 + 6 * Rnd())
            lbl_andra.Text = t2
        End If
        If Not s3 Then
            t3 = Int(1 + 6 * Rnd())
            lbl_tredje.Text = t3
        End If
        If Not s4 Then
            t4 = Int(1 + 6 * Rnd())
            lbl_fjärde.Text = t4
        End If
        If Not s5 Then
            t5 = Int(1 + 6 * Rnd())
            lbl_femte.Text = t5
        End If
        cmd_spara1.Enabled = True
        cmd_spara2.Enabled = True
        cmd_spara3.Enabled = True
        cmd_spara4.Enabled = True
        cmd_spara5.Enabled = True
        If slag = 3 Then
            cmd_slå.Enabled = False
            cmd_spara1.Enabled = False
            cmd_spara2.Enabled = False
            cmd_spara3.Enabled = False
            cmd_spara4.Enabled = False
            cmd_spara5.Enabled = False
        End If
    End Sub

    Private Sub cmd_spara1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_spara1.Click
        If s1 = False Then
            s1 = True
            cmd_spara1.Text = "Bort"
        Else
            s1 = False
            cmd_spara1.Text = "Spara"
        End If
    End Sub

    Private Sub cmd_spara2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_spara2.Click
        If s2 = False Then
            s2 = True
            cmd_spara2.Text = "Bort"
        Else
            s2 = False
            cmd_spara2.Text = "Spara"
        End If
    End Sub

    Private Sub cmd_spara3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_spara3.Click
        If s3 = False Then
            s3 = True
            cmd_spara3.Text = "Bort"
        Else
            s3 = False
            cmd_spara3.Text = "Spara"
        End If
    End Sub

    Private Sub cmd_spara4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_spara4.Click
        If s4 = False Then
            s4 = True
            cmd_spara4.Text = "Bort"
        Else
            s4 = False
            cmd_spara4.Text = "Spara"
        End If
    End Sub

    Private Sub cmd_spara5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_spara5.Click
        If s5 = False Then
            s5 = True
            cmd_spara5.Text = "Bort"
        Else
            s5 = False
            cmd_spara5.Text = "Spara"
        End If
    End Sub

    Private Sub lbl_ett_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_ett.Click
        If lbl_ett.Text = "" And slag > 0 Then
            If t1 = 1 Then
                antal = antal + 1
            End If
            If t2 = 1 Then
                antal = antal + 1
            End If
            If t3 = 1 Then
                antal = antal + 1
            End If
            If t4 = 1 Then
                antal = antal + 1
            End If
            If t5 = 1 Then
                antal = antal + 1
            End If
            lbl_ett.Text = antal
            summa = summa + antal
            totalsumma = totalsumma + antal
            Call Rensa()
        End If
    End Sub

    Private Sub lbl_två_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_två.Click
        If lbl_två.Text = "" And slag > 0 Then
            If t1 = 2 Then
                antal = antal + 1
            End If
            If t2 = 2 Then
                antal = antal + 1
            End If
            If t3 = 2 Then
                antal = antal + 1
            End If
            If t4 = 2 Then
                antal = antal + 1
            End If
            If t5 = 2 Then
                antal = antal + 1
            End If
            lbl_två.Text = antal * 2
            summa = summa + antal * 2
            totalsumma = totalsumma + antal * 2
            Call Rensa()
        End If
    End Sub

    Private Sub lbl_tre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_tre.Click
        If lbl_tre.Text = "" And slag > 0 Then
            If t1 = 3 Then
                antal = antal + 1
            End If
            If t2 = 3 Then
                antal = antal + 1
            End If
            If t3 = 3 Then
                antal = antal + 1
            End If
            If t4 = 3 Then
                antal = antal + 1
            End If
            If t5 = 3 Then
                antal = antal + 1
            End If
            lbl_tre.Text = antal * 3
            summa = summa + antal * 3
            totalsumma = totalsumma + antal * 3
            Call Rensa()
        End If
    End Sub

    Private Sub lbl_fyra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_fyra.Click
        If lbl_fyra.Text = "" And slag > 0 Then
            If t1 = 4 Then
                antal = antal + 1
            End If
            If t2 = 4 Then
                antal = antal + 1
            End If
            If t3 = 4 Then
                antal = antal + 1
            End If
            If t4 = 4 Then
                antal = antal + 1
            End If
            If t5 = 4 Then
                antal = antal + 1
            End If
            lbl_fyra.Text = antal * 4
            summa = summa + antal * 4
            totalsumma = totalsumma + antal * 4
            Call Rensa()
        End If
    End Sub

    Private Sub lbl_fem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_fem.Click
        If lbl_fem.Text = "" And slag > 0 Then
            If t1 = 5 Then
                antal = antal + 1
            End If
            If t2 = 5 Then
                antal = antal + 1
            End If
            If t3 = 5 Then
                antal = antal + 1
            End If
            If t4 = 5 Then
                antal = antal + 1
            End If
            If t5 = 5 Then
                antal = antal + 1
            End If
            lbl_fem.Text = antal * 5
            summa = summa + antal * 5
            totalsumma = totalsumma + antal * 5
            Call Rensa()
        End If
    End Sub

    Private Sub lbl_sex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_sex.Click
        If lbl_sex.Text = "" And slag > 0 Then
            If t1 = 6 Then
                antal = antal + 1
            End If
            If t2 = 6 Then
                antal = antal + 1
            End If
            If t3 = 6 Then
                antal = antal + 1
            End If
            If t4 = 6 Then
                antal = antal + 1
            End If
            If t5 = 6 Then
                antal = antal + 1
            End If
            lbl_sex.Text = antal * 6
            summa = summa + antal * 6
            totalsumma = totalsumma + antal * 6
            Call Rensa()
        End If
    End Sub

    Private Sub lbl_par_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_par.Click
        If lbl_par.Text = "" And slag > 0 Then
            For i = 6 To 1 Step -1
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal >= 2 Then
                    lbl_par.Text = i * 2
                    totalsumma = totalsumma + i * 2
                    lbl_total.Text = totalsumma
                    antal = 0
                    Exit For
                End If
                antal = 0
            Next
            If lbl_par.Text = "" Then
                lbl_par.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_tvåpar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_tvåpar.Click
        If lbl_tvåpar.Text = "" And slag > 0 Then
            For i = 6 To 1 Step -1
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal >= 2 Then
                    antal = 0
                    For i2 = i - 1 To 1 Step -1
                        If t1 = i2 Then
                            antal = antal + 1
                        End If
                        If t2 = i2 Then
                            antal = antal + 1
                        End If
                        If t3 = i2 Then
                            antal = antal + 1
                        End If
                        If t4 = i2 Then
                            antal = antal + 1
                        End If
                        If t5 = i2 Then
                            antal = antal + 1
                        End If
                        If antal >= 2 Then
                            lbl_tvåpar.Text = i * 2 + i2 * 2
                            totalsumma = totalsumma + i * 2 + i2 * 2
                            lbl_total.Text = totalsumma
                            antal = 0
                            Exit For
                        End If
                    Next
                    Exit For
                End If
                antal = 0
            Next
            If lbl_tvåpar.Text = "" Then
                lbl_tvåpar.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_triss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_triss.Click
        If lbl_triss.Text = "" And slag > 0 Then
            For i = 6 To 1 Step -1
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal >= 3 Then
                    lbl_triss.Text = i * 3
                    totalsumma = totalsumma + i * 3
                    lbl_total.Text = totalsumma
                    antal = 0
                    Exit For
                End If
                antal = 0
            Next
            If lbl_triss.Text = "" Then
                lbl_triss.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_fyrtal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_fyrtal.Click
        If lbl_fyrtal.Text = "" And slag > 0 Then
            For i = 6 To 1 Step -1
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal >= 4 Then
                    lbl_fyrtal.Text = i * 4
                    totalsumma = totalsumma + i * 4
                    lbl_total.Text = totalsumma
                    antal = 0
                    Exit For
                End If
                antal = 0
            Next
            If lbl_fyrtal.Text = "" Then
                lbl_fyrtal.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_litenstege_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_litenstege.Click
        Dim antal2 As Integer
        If lbl_litenstege.Text = "" And slag > 0 Then
            For i = 1 To 5
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal = 1 Then
                    antal = 0
                    antal2 = antal2 + 1
                    If antal2 = 5 Then
                        lbl_litenstege.Text = "15"
                        totalsumma = totalsumma + "15"
                        lbl_total.Text = totalsumma
                    End If
                Else
                    Exit For
                End If
            Next
            If lbl_litenstege.Text = "" Then
                lbl_litenstege.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_storstege_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_storstege.Click
        Dim antal2 As Integer
        If lbl_storstege.Text = "" And slag > 0 Then
            For i = 2 To 6
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal = 1 Then
                    antal = 0
                    antal2 = antal2 + 1
                    If antal2 = 5 Then
                        lbl_storstege.Text = "20"
                        totalsumma = totalsumma + "20"
                        lbl_total.Text = totalsumma
                    End If
                Else
                    Exit For
                End If
            Next
            If lbl_storstege.Text = "" Then
                lbl_storstege.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_chans_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_chans.Click
        If lbl_chans.Text = "" And slag > 0 Then
            lbl_chans.Text = t1 + t2 + t3 + t4 + t5
            totalsumma = totalsumma + t1 + t2 + t3 + t4 + t5
            lbl_total.Text = totalsumma
        End If
        Call Rensa2()
    End Sub

    Private Sub lbl_kåk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_kåk.Click
        If lbl_kåk.Text = "" And slag > 0 Then
            For i = 6 To 1 Step -1
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal = 3 Then
                    antal = 0
                    For i2 = 6 To 1 Step -1
                        If t1 = i2 Then
                            antal = antal + 1
                        End If
                        If t2 = i2 Then
                            antal = antal + 1
                        End If
                        If t3 = i2 Then
                            antal = antal + 1
                        End If
                        If t4 = i2 Then
                            antal = antal + 1
                        End If
                        If t5 = i2 Then
                            antal = antal + 1
                        End If
                        If antal = 2 Then
                            lbl_kåk.Text = i * 3 + i2 * 2
                            totalsumma = totalsumma + i * 3 + i2 * 2
                            lbl_total.Text = totalsumma
                            antal = 0
                            Exit For
                        End If
                        antal = 0
                    Next
                    Exit For
                End If
                antal = 0
            Next
            If lbl_kåk.Text = "" Then
                lbl_kåk.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub

    Private Sub lbl_yatzy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_yatzy.Click
        If lbl_yatzy.Text = "" And slag > 0 Then
            For i = 1 To 6
                If t1 = i Then
                    antal = antal + 1
                End If
                If t2 = i Then
                    antal = antal + 1
                End If
                If t3 = i Then
                    antal = antal + 1
                End If
                If t4 = i Then
                    antal = antal + 1
                End If
                If t5 = i Then
                    antal = antal + 1
                End If
                If antal = 5 Then
                    lbl_yatzy.Text = "50"
                    totalsumma = totalsumma + 50
                    lbl_total.Text = totalsumma
                    antal = 0
                    Exit For
                End If
                antal = 0
            Next
            If lbl_yatzy.Text = "" Then
                lbl_yatzy.Text = "0"
            End If
            Call Rensa2()
        End If
    End Sub
End Class
