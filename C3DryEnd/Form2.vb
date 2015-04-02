Public Class FRM_Menu


    Private Sub ShowToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ShowToolStripMenuItem3.Click
        D3_Dry.Visible = True
    End Sub


    Private Sub HideToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles HideToolStripMenuItem3.Click
        D3_Dry.Visible = False
    End Sub

    

    Private Sub ShowToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ShowToolStripMenuItem4.Click
        FrmD8_Dry.Visible = True
    End Sub


    Private Sub HideToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles HideToolStripMenuItem4.Click
        FrmD8_Dry.Visible = False
    End Sub
End Class