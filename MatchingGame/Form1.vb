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


    ' firstClickedは、最初にクリックしたラベルを指します。 
    Private firstClicked As Label = Nothing

    ' secondClickedは2番目にクリックしたラベルを指します。
    Private secondClicked As Label = Nothing

    Private Sub label_Click(ByVal sender As System.Object,
                        ByVal e As System.EventArgs) Handles Label1.Click,
    Label2.Click, Label3.Click, Label4.Click, Label5.Click, Label6.Click,
    Label7.Click, Label8.Click, Label9.Click, Label10.Click, Label11.Click,
    Label12.Click, Label13.Click, Label14.Click, Label15.Click, Label16.Click

        ' マッチしないアイコンが2つ表示されてる場合、タイマーが作動しているため、
        ' タイマー作動中はクリック動作を無視します。
        If Timer1.Enabled Then Exit Sub

        Dim clickedLabel = TryCast(sender, Label)

        If clickedLabel IsNot Nothing Then

            ' クリックされたラベルの色が黒の場合、
            ' 既にクリックされている為、クリックを無視する。 
            If clickedLabel.ForeColor = Color.Black Then Exit Sub

            ' firstClickedがNothingの場合、これはプレーヤーがクリックしたペアの最初のアイコンなので、
            ' プレーヤーがクリックしたラベルにfirstClickedを設定し、その色を黒に変更してreturnを返します。
            If firstClicked Is Nothing Then
                firstClicked = clickedLabel
                firstClicked.ForeColor = Color.Black
                Exit Sub
            End If

            ' ここが動作する際はタイマーが作動しておらず、firstClickedはNothingではないため、
            ' プレーヤーがクリックしたペアの2番目のアイコンとなります。
            ' プレーヤーがクリックしたラベルにsecondClickedを設定し、その色を黒に変更してreturnを返します。
            secondClicked = clickedLabel
            secondClicked.ForeColor = Color.Black

            ' プレーヤーがクリア条件を満たしているか確認します。
            CheckForWinner()

            ' プレイヤーがクリックしたアイコンが一致している場合、
            ' それらのアイコンは黒色のままにしてfirstClickedとsecondClickedをリセットします。
            If firstClicked.Text = secondClicked.Text Then
                firstClicked = Nothing
                secondClicked = Nothing
                Exit Sub
            End If

            ' ここが動作する際はプレイヤーが2つの異なるアイコンをクリックしているため、
            ' タイマーをスタートさせる。(タイマー終了時にアイコンを隠してfirstClickedとsecondClickedをリセットする。)
            Timer1.Start()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' タイマーを止める。
        Timer1.Stop()

        ' アイコンを背景色と同じ色に戻す。
        firstClicked.ForeColor = firstClicked.BackColor
        secondClicked.ForeColor = secondClicked.BackColor

        ' firstClickedとsecondClickedをリセットする。
        firstClicked = Nothing
        secondClicked = Nothing
    End Sub

    Private Sub CheckForWinner()

        ' TableLayoutPanel内のすべてのラベルを調べます、
        ' それぞれのアイコンで背景と文字色が同じかをチェックする。
        For Each control In TableLayoutPanel1.Controls
            Dim iconLabel = TryCast(control, Label)
            If iconLabel IsNot Nothing AndAlso
               iconLabel.ForeColor = iconLabel.BackColor Then Exit Sub
        Next

        ' 同じ物がない場合は、すべてのアイコンを一致させられているため、クリア処理を行う。
        ' メッセージを表示してフォームを閉じます。
        MessageBox.Show("全てのアイコンを一致させました！", "Congratulations")
        Close()

    End Sub

End Class
