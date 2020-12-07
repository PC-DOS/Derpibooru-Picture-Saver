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
    Dim iPageCount As Integer
    Dim iPauseThreshold As Integer
    Dim iPauseDuration As Integer
    Dim UserSpecifiedFilterID As Integer
    Dim CurrentSearchPrefix As String = "https://derpibooru.org/api/v1/json/search?q="
    Const DerpibooruSearchPrefix As String = "https://derpibooru.org/api/v1/json/search?q="
    Const DerpibooruSearchPrefixBackup As String = "https://trixiebooru.org/api/v1/json/search?q="
    Const DerpibooruSearchPageSelector As String = "&page="
    Const DerpibooruImagesPerpageSelector As String = "&per_page=50"
    Const DerpibooruImagesPerpage As Integer = 50
    Const DerpibooruImagesMinScoreSelector As String = "&min_score="
    Const DerpibooruImagesMaxScoreSelector As String = "&max_score="
    Const DerpibooruImagesFilterSelector As String = "&filter_id="
    Const DerpibooruImagesSortFieldSelector As String = "&sf="
    Const DerpibooruImagesSortDirectionSelector As String = "&sq="
    Private Sub RefreshURLList()
        lstSavedURL.ItemsSource = EmptyList
        lstSavedURL.ItemsSource = URLList
        lstSavedURL.SelectedIndex = lstSavedURL.Items.Count - 1
        Try
            lstSavedURL.ScrollIntoView(lstSavedURL.SelectedItem)
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub LockWindow()
        txtMaxScore.IsEnabled = False
        txtMinScore.IsEnabled = False
        txtPauseDuration.IsEnabled = False
        txtPauseThreshold.IsEnabled = False
        txtSaveTo.IsEnabled = False
        txtSearchKey.IsEnabled = False
        txtPageCount.IsEnabled = False
        btnBrowse.IsEnabled = False
        btnStart.IsEnabled = False
        chkPause.IsEnabled = False
        chkRestrictMaxScore.IsEnabled = False
        chkRestrictMinScore.IsEnabled = False
        chkRestrictPageCount.IsEnabled = False
        chkRestrictMinWilsonScore.IsEnabled = False
        chkUseTrixieBooru.IsEnabled = False
        cmbFilters.IsEnabled = False
        cmbSordField.IsEnabled = False
        cmbSortDirection.IsEnabled = False
        sldMinWilsonScore.IsEnabled = False
    End Sub
    Private Sub UnlockWindow()
        txtMaxScore.IsEnabled = chkRestrictMaxScore.IsChecked
        txtMinScore.IsEnabled = chkRestrictMinScore.IsChecked
        txtPauseDuration.IsEnabled = chkPause.IsChecked
        txtPauseThreshold.IsEnabled = chkPause.IsChecked
        txtSaveTo.IsEnabled = True
        txtSearchKey.IsEnabled = True
        txtPageCount.IsEnabled = chkRestrictPageCount.IsChecked
        btnBrowse.IsEnabled = True
        btnStart.IsEnabled = True
        chkPause.IsEnabled = True
        chkRestrictMaxScore.IsEnabled = True
        chkRestrictMinScore.IsEnabled = True
        chkRestrictPageCount.IsEnabled = True
        chkRestrictMinWilsonScore.IsEnabled = True
        chkUseTrixieBooru.IsEnabled = True
        cmbFilters.IsEnabled = True
        cmbSordField.IsEnabled = True
        cmbSortDirection.IsEnabled = True
        sldMinWilsonScore.IsEnabled = chkRestrictMinWilsonScore.IsChecked
    End Sub
    Private Sub SetTaskbarProgess(MaxValue As Integer, MinValue As Integer, CurrentValue As Integer, Optional State As Shell.TaskbarItemProgressState = Shell.TaskbarItemProgressState.Normal)
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
            .Description = "請指定下載的檔案的儲存位置，然後按一下 [確定] 按鈕。"
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
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End If
        If chkPause.IsChecked Then
            If Not IsNumeric(txtPauseDuration.Text) Then
                MessageBox.Show("指定的擱置時間長度存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Not IsNumeric(txtPauseThreshold.Text) Then
                MessageBox.Show("指定的擱置門限值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Int(txtPauseDuration.Text) <= 0 Then
                MessageBox.Show("指定的擱置時間長度必須為正數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Int(txtPauseThreshold.Text) <= 0 Then
                MessageBox.Show("指定的擱置門限值必須為正數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iPauseThreshold = Math.Abs(Int(txtPauseThreshold.Text))
            iPauseDuration = Math.Abs(Int(txtPauseDuration.Text))
        End If
        If chkRestrictPageCount.IsChecked Then
            If Not IsNumeric(txtPageCount.Text) Then
                MessageBox.Show("指定的下載頁數存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Int(txtPageCount.Text) < 1 Then
                MessageBox.Show("指定的下載頁數必須為正數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iPageCount = txtPageCount.Text
        End If
        Dim iPageIndex As Integer = 1
        Dim Filter As New ComboBoxItem
        Dim SortField As ComboBoxItem = cmbSordField.SelectedItem
        Dim SortDirection As ComboBoxItem = cmbSortDirection.SelectedItem
        Filter = cmbFilters.SelectedItem
        sSearchQuery = CurrentSearchPrefix & txtSearchKey.Text.Trim().Replace(" ", "+") & _
                       DerpibooruSearchPageSelector & iPageIndex.ToString() & _
                       DerpibooruImagesPerpageSelector & _
                       DerpibooruImagesSortFieldSelector & SortField.Tag.ToString & _
                       DerpibooruImagesSortDirectionSelector & SortDirection.Tag.ToString & _
                       DerpibooruImagesFilterSelector & IIf(Filter.Tag = -1, UserSpecifiedFilterID.ToString, Filter.Tag.ToString)
        If chkRestrictMinScore.IsChecked Then
            If Not IsNumeric(txtMinScore.Text) Then
                MessageBox.Show("指定的最低評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iMinScore = txtMinScore.Text
            'sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString()
        End If
        If chkRestrictMaxScore.IsChecked Then
            If Not IsNumeric(txtMaxScore.Text) Then
                MessageBox.Show("指定的最高評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            iMaxScore = txtMaxScore.Text
            'sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString()
        End If
        If chkRestrictMaxScore.IsChecked And chkRestrictMinScore.IsChecked Then
            If iMinScore > iMaxScore Then
                If MessageBox.Show("指定的最低評分值大於指定的最高評分值，這可能導致例外情況，您確定要繼續嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Forms.DialogResult.Yes Then
                    'DoNothing()
                Else
                    UnlockWindow()
                    SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
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
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
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
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
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
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End Try
            JSONResponse = JsonConvert.DeserializeObject(sJSON)
            iImageTotal = JSONResponse("total")
            iPageTotal = Math.Ceiling((CInt(JSONResponse("total")) / 50))
            If chkRestrictPageCount.IsChecked Then
                If iPageTotal > iPageCount Then
                    iPageTotal = iPageCount
                    iImageTotal = 50 * iPageCount
                End If
            End If
            prgProgress.Maximum = iPageTotal
            prgProgress.Minimum = 0
            prgProgress.Value = 0
            For iPageIndex = 1 To iPageTotal
                sSearchQuery = CurrentSearchPrefix & txtSearchKey.Text.Trim().Replace(" ", "+") & _
                               DerpibooruSearchPageSelector & iPageIndex.ToString() & _
                               DerpibooruImagesPerpageSelector & _
                               DerpibooruImagesSortFieldSelector & SortField.Tag.ToString & _
                               DerpibooruImagesSortDirectionSelector & SortDirection.Tag.ToString & _
                               DerpibooruImagesFilterSelector & IIf(Filter.Tag = -1, UserSpecifiedFilterID.ToString, Filter.Tag.ToString)
                If chkRestrictMinScore.IsChecked Then
                    If Not IsNumeric(txtMinScore.Text) Then
                        MessageBox.Show("指定的最低評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UnlockWindow()
                        SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                        Exit Sub
                    End If
                    iMinScore = txtMinScore.Text
                    'sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString() '由於 API 變更，已廢除。
                End If
                If chkRestrictMaxScore.IsChecked Then
                    If Not IsNumeric(txtMaxScore.Text) Then
                        MessageBox.Show("指定的最高評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UnlockWindow()
                        SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                        Exit Sub
                    End If
                    iMaxScore = txtMaxScore.Text
                    'sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString() '由於 API 變更，已廢除。
                End If
                If chkRestrictMaxScore.IsChecked And chkRestrictMinScore.IsChecked Then
                    If iMinScore > iMaxScore Then
                        If MessageBox.Show("指定的最低評分值大於指定的最高評分值，可能導致例外情況，您確定要繼續嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Forms.DialogResult.Yes Then
                            'DoNothing()
                        Else
                            UnlockWindow()
                            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
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
                    SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                    Exit Sub
                End Try
                JSONCache.Add(sJSON)
                System.Windows.Forms.Application.DoEvents()
                prgProgress.Value = iPageIndex
                SetTaskbarProgess(iPageTotal, 0, iPageIndex, Shell.TaskbarItemProgressState.Paused)
            Next
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            txtStatus.Text = "正在從 Derpibooru 下載檔案。"
            sJSON = JSONCache(0)
            JSONResponse = JsonConvert.DeserializeObject(sJSON)
            iImageTotal = JSONResponse("total")
            iPageTotal = Math.Ceiling((CInt(JSONResponse("total")) / 50))
            If chkRestrictPageCount.IsChecked Then
                If iPageTotal > iPageCount Then
                    iPageTotal = iPageCount
                    iImageTotal = 50 * iPageCount
                End If
            End If
            prgProgress.Maximum = iImageTotal
            prgProgress.Minimum = 0
            prgProgress.Value = 0
            Dim sImageFileName As String = ""
            Dim sImageURL As String = ""
            Dim nSuccess As Integer = 0
            Dim nFail As Integer = 0
            Dim nIgnored As Integer = 0
            For iPageIndex = 1 To iPageTotal
                sJSON = JSONCache(iPageIndex - 1)
                JSONResponse = JsonConvert.DeserializeObject(sJSON)
                For iPhotoIndex = 0 To IIf(iPageIndex = iPageTotal, iImageTotal - (iPageIndex - 1) * 50 - 1, 49)
                    If Not Directory.Exists(sSaveTo) Then
                        Directory.CreateDirectory(sSaveTo)
                    End If
                    sImageURL = JSONResponse("images")(iPhotoIndex)("view_url").ToString
                    sImageFileName = GetFileNameFromDircectURL(sImageURL)
                    'MessageBox.Show(sImageFileName)
                    If chkRestrictMinScore.IsChecked Then
                        If CInt(JSONResponse("images")(iPhotoIndex)("score")) < iMinScore Then
                            nIgnored += 1
                            URLList.Add("已略過 " & sImageURL & " 的下載作業。因為其評分 " & JSONResponse("images")(iPhotoIndex)("score").ToString() & " 低於指定的最低評分 " & iMinScore.ToString() & "。")
                            RefreshURLList()
                            System.Windows.Forms.Application.DoEvents()
                            prgProgress.Value += 1
                            SetTaskbarProgess(iImageTotal, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                            UpdateLayout()
                            If chkPause.IsChecked Then
                                If prgProgress.Value Mod iPauseThreshold = 0 Then
                                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(iPauseDuration))
                                End If
                            End If
                            Continue For
                        End If
                    End If
                    If chkRestrictMaxScore.IsChecked Then
                        If CInt(JSONResponse("images")(iPhotoIndex)("score")) > iMaxScore Then
                            nIgnored += 1
                            URLList.Add("已略過 " & sImageURL & " 的下載作業。因為其評分 " & JSONResponse("images")(iPhotoIndex)("score").ToString() & " 高於指定的最高評分 " & iMaxScore.ToString() & "。")
                            RefreshURLList()
                            System.Windows.Forms.Application.DoEvents()
                            prgProgress.Value += 1
                            SetTaskbarProgess(iImageTotal, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                            UpdateLayout()
                            If chkPause.IsChecked Then
                                If prgProgress.Value Mod iPauseThreshold = 0 Then
                                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(iPauseDuration))
                                End If
                            End If
                            Continue For
                        End If
                    End If
                    If chkRestrictMinWilsonScore.IsChecked Then
                        If CDbl(JSONResponse("images")(iPhotoIndex)("wilson_score")) < sldMinWilsonScore.Value Then
                            nIgnored += 1
                            URLList.Add("已略過 " & sImageURL & " 的下載作業。因為其質量評分 " & JSONResponse("images")(iPhotoIndex)("wilson_score").ToString() & " 低於指定的最低質量評分 " & sldMinWilsonScore.Value.ToString() & "。")
                            RefreshURLList()
                            System.Windows.Forms.Application.DoEvents()
                            prgProgress.Value += 1
                            SetTaskbarProgess(iImageTotal, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                            UpdateLayout()
                            If chkPause.IsChecked Then
                                If prgProgress.Value Mod iPauseThreshold = 0 Then
                                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(iPauseDuration))
                                End If
                            End If
                            Continue For
                        End If
                    End If
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
                    SetTaskbarProgess(iImageTotal, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                    UpdateLayout()
                    If chkPause.IsChecked Then
                        If prgProgress.Value Mod iPauseThreshold = 0 Then
                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(iPauseDuration))
                        End If
                    End If
                Next
            Next
            MessageBox.Show("作業成功完成。本次共成功下載了 " & nSuccess.ToString & " 個檔案，有 " & nFail.ToString & " 個檔案下載失敗，略過了 " & nIgnored.ToString() & " 個檔案。", "大功告成!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            UnlockWindow()
            txtStatus.Text = "作業成功完成。本次共成功下載了 " & nSuccess.ToString & " 個檔案，有 " & nFail.ToString & " 個檔案下載失敗，略過了 " & nIgnored.ToString() & " 個檔案。"
            UpdateLayout()
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
        Else
            UnlockWindow()
            txtStatus.Text = "使用者取消了作業。"
            UpdateLayout()
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
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
        Try
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurituProctocolTypeExtensions.Tls11 Or SecurituProctocolTypeExtensions.Tls12
        Catch ex As Exception
            MessageBox.Show("無法配置應用程式以啟用對 TLS 1.1 和 TLS 1.2 的支援，因為發生例外情況:" & vbCrLf & ex.Message & vbCrLf & vbCrLf & "如果您正在使用 Windows 7 或更早版本的 Windows 作業系統，那麼您可能需要更新您的作業系統。" & vbCrLf & "應用程式仍將繼續啟動，但是可能無法正常使用。", "無法啟用必要的協力組件", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        If chkUseTrixieBooru.IsChecked Then
            CurrentSearchPrefix = DerpibooruSearchPrefixBackup
        Else
            CurrentSearchPrefix = DerpibooruSearchPrefix
        End If
        txtSearchKey.Focus()
    End Sub

    Private Sub chkRestrictPageCount_Click(sender As Object, e As RoutedEventArgs) Handles chkRestrictPageCount.Click
        txtPageCount.IsEnabled = chkRestrictPageCount.IsChecked
    End Sub

    Private Sub cmbFilters_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbFilters.SelectionChanged
        If cmbFilters.SelectedIndex = 6 Then
            Try
                UserSpecifiedFilterID = Int(InputBox("請指定您所使用的過濾器的 ID 編號。", "自訂過濾器", 0))
            Catch ex As Exception
                UserSpecifiedFilterID = 0
            End Try
            MessageBox.Show("目前使用的過濾器 ID 編號為 " & UserSpecifiedFilterID & "。", "自訂過濾器", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub chkRestrictMinWilsonScore_Click(sender As Object, e As RoutedEventArgs) Handles chkRestrictMinWilsonScore.Click
        sldMinWilsonScore.IsEnabled = chkRestrictMinWilsonScore.IsChecked
    End Sub

    Private Sub chkUseTrixieBooru_Click(sender As Object, e As RoutedEventArgs) Handles chkUseTrixieBooru.Click
        If chkUseTrixieBooru.IsChecked Then
            CurrentSearchPrefix = DerpibooruSearchPrefixBackup
        Else
            CurrentSearchPrefix = DerpibooruSearchPrefix
        End If
    End Sub
End Class
