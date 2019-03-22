Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Window
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Class MainWindow
    Dim sSaveTo As String = "C:\Derpibooru Images\"
    Dim sSearchQuery As String
    Dim iMinScore As Integer
    Dim iMaxScore As Integer
    Dim URLList As New List(Of String)
    Dim EmptyList As New List(Of String)
    Dim JSONCache As New List(Of String)
    Dim IsPauseOn As Boolean = False
    Dim iPauseThreshold As Integer
    Dim iPauseDuration As Integer
    Const DerpibooruSearchPreix As String = "https://derpibooru.org/search.json?q="
    Const DerpibooruSearchPageSelector As String = "&page="
    Const DerpibooruImagesPerpageSelector As String = "&perpage=50"
    Const DerpibooruImagesPerpage As Integer = 50
    Const DerpibooruImagesMinScoreSelector As String = "&min_score="
    Const DerpibooruImagesMaxScoreSelector As String = "&max_score="
    Const DerpibooruImagesFilterSelector As String = "&filter_id="
    Private Sub RefreshURLList()
        lstSavedURL.ItemsSource = EmptyList
        lstSavedURL.ItemsSource = URLList
    End Sub
    Private Sub LockWindow()
        txtMaxScore.IsEnabled = False
        txtMinScore.IsEnabled = False
        txtPauseDuration.IsEnabled = False
        txtPauseThreshold.IsEnabled = False
        txtSaveTo.IsEnabled = False
        txtSearchKey.IsEnabled = False
        btnBrowse.IsEnabled = False
        btnStart.IsEnabled = False
        chkPause.IsEnabled = False
        chkRestrictMaxScore.IsEnabled = False
        chkRestrictMinScore.IsEnabled = False
        cmbFilters.IsEnabled = False
    End Sub
    Private Sub UnlockWindow()
        txtMaxScore.IsEnabled = chkRestrictMaxScore.IsChecked
        txtMinScore.IsEnabled = chkRestrictMinScore.IsChecked
        txtPauseDuration.IsEnabled = chkPause.IsChecked
        txtPauseThreshold.IsEnabled = chkPause.IsChecked
        txtSaveTo.IsEnabled = True
        txtSearchKey.IsEnabled = True
        btnBrowse.IsEnabled = True
        btnStart.IsEnabled = True
        chkPause.IsEnabled = True
        chkRestrictMaxScore.IsEnabled = True
        chkRestrictMinScore.IsEnabled = True
        cmbFilters.IsEnabled = False
    End Sub
    Private Sub SetTastbarProgess(MaxValue As Integer, MinValue As Integer, CurrentValue As Integer, Optional State As Shell.TaskbarItemProgressState = Shell.TaskbarItemProgressState.Normal)
        If MaxValue <= MinValue Or CurrentValue < MinValue Or CurrentValue > MaxValue Then
            Exit Sub
        End If
        TaskbarItem.ProgressValue = (CurrentValue - MinValue) / (MaxValue - MinValue)
        TaskbarItem.ProgressState = State
    End Sub
    Private Function GetFileNameFromDircectURL(URL As String) As String
        If URL.Trim = "" Then
            Return ""
        End If
        URL = Trim(URL)
        If URL(URL.Length - 1) = "/" Then
            Return ""
        End If
        Dim i As Integer
        For i = URL.Length - 1 To 0 Step -1
            If URL(i) = "/" Then
                Exit For
            End If
        Next
        Return URL.Substring(i + 1)
    End Function
    Private Sub btnBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowse.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "請指定儲存位置"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            sSaveTo = FolderBrowser.SelectedPath
            If sSaveTo(sSaveTo.Length - 1) <> "\" Then
                sSaveTo = sSaveTo & "\"
            End If
            txtSaveTo.Text = sSaveTo
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As RoutedEventArgs) Handles btnStart.Click
        LockWindow()
        prgProgress.Minimum = 0
        prgProgress.Maximum = 100
        prgProgress.Value = 0
        If txtSearchKey.Text.Trim() = "" Then
            MessageBox.Show("請輸入搜尋條件。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockWindow()
            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End If
        If chkPause.IsChecked Then
            If Not IsNumeric(txtPauseDuration.Text) Then
                MessageBox.Show("指定的擱置時間長度存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Not IsNumeric(txtPauseThreshold.Text) Then
                MessageBox.Show("指定的擱置門限值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iPauseThreshold = txtPauseThreshold.Text
            iPauseDuration = txtPauseDuration.Text
            IsPauseOn = True
        Else
            IsPauseOn = False
        End If
        Dim iPageIndex As Integer = 1
        Dim Filter As New ComboBoxItem
        Filter = cmbFilters.SelectedItem
        sSearchQuery = DerpibooruSearchPreix & txtSearchKey.Text.Trim().Replace(" ", "+") & DerpibooruSearchPageSelector & iPageIndex.ToString() & DerpibooruImagesPerpageSelector & DerpibooruImagesFilterSelector & Filter.Tag.ToString
        If chkRestrictMinScore.IsChecked Then
            If Not IsNumeric(txtMinScore.Text) Then
                MessageBox.Show("指定的最低評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iMinScore = txtMinScore.Text
            sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString()
        End If
        If chkRestrictMaxScore.IsChecked Then
            If Not IsNumeric(txtMaxScore.Text) Then
                MessageBox.Show("指定的最高評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iMaxScore = txtMaxScore.Text
            sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString()
        End If
        If chkRestrictMaxScore.IsChecked And chkRestrictMinScore.IsChecked Then
            If iMinScore > iMaxScore Then
                If MessageBox.Show("指定的最低評分值大於指定的最高評分值，這可能導致例外情況，您確定要繼續嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Forms.DialogResult.Yes Then
                    'DoNothing()
                Else
                    UnlockWindow()
                    SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                    Exit Sub
                End If
            End If
        End If
        'MessageBox.Show(sSearchQuery)
        URLList.Clear()
        RefreshURLList()
        JSONCache.Clear()
        txtStatus.Text = "正在連線到 " & sSearchQuery & " 以 GET 方法擷取 JSON 檔案。"
        UpdateLayout()
        Dim sJSON As String = ""
        Dim iPageTotal As Integer
        Dim HttpReq As New SimpleHttpRequest.HttpJsonRequest
        Try
            sJSON = HttpReq.DoGet(sSearchQuery)
        Catch ex As Exception
            MessageBox.Show("發生例外情況:" & vbCrLf & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockWindow()
            txtStatus.Text = "發生例外情況: " & ex.Message & "，結束作業。"
            UpdateLayout()
            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End Try
        'MessageBox.Show(sJSON)
        Dim JSONResponse As New JObject
        JSONResponse = JsonConvert.DeserializeObject(sJSON)
        'MessageBox.Show(JSONResponse("total").ToString)
        Dim iImageTotal As Integer = JSONResponse("total")
        iPageTotal = Math.Ceiling((CInt(JSONResponse("total")) / 50))
        txtStatus.Text = "JSON 檔案擷取完畢，總共搜尋到 " & JSONResponse("total").ToString & " 張相片。一共 " & iPageTotal.ToString & " 個分頁。"
        UpdateLayout()
        If JSONResponse("total").ToString = "0" Then
            MessageBox.Show("沒有找到符合您指定的搜尋條件的相片。", "沒有結果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            UnlockWindow()
            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End If
        If MessageBox.Show("搜尋完畢!" & vbCrLf & "本次搜尋總共搜尋到 " & JSONResponse("total").ToString & " 張相片。一共 " & iPageTotal.ToString & " 個分頁。" & vbCrLf & "開始下載相片嗎?", "下載相片", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Forms.DialogResult.Yes Then
            Dim iPhotoIndex As Integer
            txtStatus.Text = "正在快取搜尋結果。"
            UpdateLayout()
            Try
                sJSON = HttpReq.DoGet(sSearchQuery)
            Catch ex As Exception
                MessageBox.Show("發生例外情況:" & vbCrLf & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtStatus.Text = "發生例外情況: " & ex.Message & "，結束作業。"
                UnlockWindow()
                UpdateLayout()
                SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End Try
            JSONResponse = JsonConvert.DeserializeObject(sJSON)
            iImageTotal = JSONResponse("total")
            iPageTotal = Math.Ceiling((CInt(JSONResponse("total")) / 50))
            prgProgress.Maximum = iPageTotal
            prgProgress.Minimum = 0
            prgProgress.Value = 0
            For iPageIndex = 1 To iPageTotal
                sSearchQuery = DerpibooruSearchPreix & txtSearchKey.Text.Trim().Replace(" ", "+") & DerpibooruSearchPageSelector & iPageIndex.ToString() & DerpibooruImagesPerpageSelector & DerpibooruImagesFilterSelector & Filter.Tag.ToString
                If chkRestrictMinScore.IsChecked Then
                    If Not IsNumeric(txtMinScore.Text) Then
                        MessageBox.Show("指定的最低評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UnlockWindow()
                        SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                        Exit Sub
                    End If
                    iMinScore = txtMinScore.Text
                    sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString()
                End If
                If chkRestrictMaxScore.IsChecked Then
                    If Not IsNumeric(txtMaxScore.Text) Then
                        MessageBox.Show("指定的最高評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UnlockWindow()
                        SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                        Exit Sub
                    End If
                    iMaxScore = txtMaxScore.Text
                    sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString()
                End If
                If chkRestrictMaxScore.IsChecked And chkRestrictMinScore.IsChecked Then
                    If iMinScore > iMaxScore Then
                        If MessageBox.Show("指定的最低評分值大於指定的最高評分值，可能導致例外情況，您確定要繼續嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Forms.DialogResult.Yes Then
                            'DoNothing()
                        Else
                            UnlockWindow()
                            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                            Exit Sub
                        End If
                    End If
                End If
                Try
                    sJSON = HttpReq.DoGet(sSearchQuery)
                Catch ex As Exception
                    MessageBox.Show("發生例外情況:" & vbCrLf & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    UnlockWindow()
                    txtStatus.Text = "發生例外情況: " & ex.Message & "，結束作業。"
                    UpdateLayout()
                    SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                    Exit Sub
                End Try
                JSONCache.Add(sJSON)
                System.Windows.Forms.Application.DoEvents()
                prgProgress.Value = iPageIndex
                SetTastbarProgess(iPageTotal, 0, iPageIndex, Shell.TaskbarItemProgressState.Paused)
            Next
            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            txtStatus.Text = "正在從 Derpibooru 下載檔案。"
            sJSON = JSONCache(0)
            JSONResponse = JsonConvert.DeserializeObject(sJSON)
            iImageTotal = JSONResponse("total")
            iPageTotal = Math.Ceiling((CInt(JSONResponse("total")) / 50))
            prgProgress.Maximum = iImageTotal
            prgProgress.Minimum = 0
            prgProgress.Value = 0
            Dim sImageFileName As String = ""
            Dim sImageURL As String = ""
            Dim nSuccess As Integer = 0
            Dim nFail As Integer = 0
            For iPageIndex = 1 To iPageTotal
                sJSON = JSONCache(iPageIndex - 1)
                JSONResponse = JsonConvert.DeserializeObject(sJSON)
                For iPhotoIndex = 0 To IIf(iPageIndex = iPageTotal, iImageTotal - (iPageIndex - 1) * 50 - 1, 49)
                    If Not Directory.Exists(sSaveTo) Then
                        Directory.CreateDirectory(sSaveTo)
                    End If
                    sImageURL = "https:" & JSONResponse("search")(iPhotoIndex)("image").ToString
                    sImageFileName = GetFileNameFromDircectURL(sImageURL)
                    Dim FileDownloader As New WebClient
                    Try
                        FileDownloader.DownloadFile(sImageURL, sSaveTo & sImageFileName)
                        nSuccess += 1
                        URLList.Add("成功從 " & sImageURL & " 下載相片到 " & sSaveTo & sImageFileName)
                        RefreshURLList()
                    Catch ex As Exception
                        URLList.Add("從 " & sImageURL & " 下載相片時失敗，發生例外情況: " & ex.Message)
                        RefreshURLList()
                        nFail += 1
                    End Try
                    System.Windows.Forms.Application.DoEvents()
                    prgProgress.Value += 1
                    SetTastbarProgess(iImageTotal, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                    UpdateLayout()
                    If IsPauseOn Then
                        If prgProgress.Value Mod iPauseThreshold = 0 Then
                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(iPauseDuration))
                        End If
                    End If
                Next
            Next
            MessageBox.Show("作業成功完成。本次共成功下載了 " & nSuccess.ToString & " 個檔案，有 " & nFail.ToString & " 個檔案下載失敗。", "大功告成!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            UnlockWindow()
            txtStatus.Text = "作業成功完成。本次共成功下載了 " & nSuccess.ToString & " 個檔案，有 " & nFail.ToString & " 個檔案下載失敗。"
            UpdateLayout()
            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
        Else
            UnlockWindow()
            txtStatus.Text = "使用者取消了作業。"
            UpdateLayout()
            SetTastbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End If
    End Sub

    Private Sub chkRestrictMaxScore_Click(sender As Object, e As RoutedEventArgs) Handles chkRestrictMaxScore.Click
        txtMaxScore.IsEnabled = chkRestrictMaxScore.IsChecked
    End Sub

    Private Sub chkRestrictMinScore_Click(sender As Object, e As RoutedEventArgs) Handles chkRestrictMinScore.Click
        txtMinScore.IsEnabled = chkRestrictMinScore.IsChecked
    End Sub

    Private Sub chkPause_Click(sender As Object, e As RoutedEventArgs) Handles chkPause.Click
        txtPauseThreshold.IsEnabled = chkPause.IsChecked
        txtPauseDuration.IsEnabled = chkPause.IsChecked
    End Sub

    Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        End
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        txtSearchKey.Focus()
    End Sub
End Class
