Public Class Form1

    ' Randomオブジェクトを使って、マスのアイコンをランダムに選択する。
    Private random As New Random

    ' これらの文字はそれぞれ、Webdingsフォントによりアイコンとして表示される。
    ' 各2つずつリストに格納する。
    Private icons =
      New List(Of String) From {"!", "!", "N", "N", ",", ",", "k", "k",
                                "b", "b", "v", "v", "w", "w", "z", "z"}

    Private Sub AssignIconsToSquares()

        ' TableLayoutPanelには16個のラベルがあり、
        ' iconsのリストには16個のアイコンがあります。
        For Each control In TableLayoutPanel1.Controls
            Dim iconLabel = TryCast(control, Label)
            If iconLabel IsNot Nothing Then
                Dim randomNumber = random.Next(icons.Count)
                iconLabel.Text = icons(randomNumber)
                iconLabel.ForeColor = iconLabel.BackColor
                icons.RemoveAt(randomNumber)
            End If
        Next

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AssignIconsToSquares()
    End Sub

    Private Sub label_Click(ByVal sender As System.Object,
                        ByVal e As System.EventArgs) Handles Label1.Click,
    Label2.Click, Label3.Click, Label4.Click, Label5.Click, Label6.Click,
    Label7.Click, Label8.Click, Label9.Click, Label10.Click, Label11.Click,
    Label12.Click, Label13.Click, Label14.Click, Label15.Click, Label16.Click

        Dim clickedLabel = TryCast(sender, Label)

        If clickedLabel IsNot Nothing Then

            ' クリックされたラベルの色が黒の場合、
            ' 既にクリックされている為、クリックを無視する。 
            If clickedLabel.ForeColor = Color.Black Then Exit Sub

            clickedLabel.ForeColor = Color.Black
        End If
    End Sub

End Class
