Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Window
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Class MainWindow
    Dim DownloadedFileSavePath As String = "C:\Derpibooru Images\"
    Dim SearchQuery As String
    Dim MinScore As Integer
    Dim MaxScore As Integer
    Dim URLList As New List(Of String)
    Dim EmptyList As New List(Of String)
    Dim JSONCache As New List(Of String)
    Dim PageIndexBegin As Integer
    Dim PageIndexEnd As Integer
    Dim PageCount As Integer
    Dim PauseThreshold As Integer
    Dim PauseDuration As Integer
    Dim UserSpecifiedFilterID As Integer
    Dim ThumbnailWidth As Integer
    Dim ThumbnailHeight As Integer
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
    Const DerpibooruImagesSortDirectionSelector As String = "&sd="
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
        txtPageIndexBegin.IsEnabled = False
        txtPageIndexEnd.IsEnabled = False
        btnBrowse.IsEnabled = False
        btnStart.IsEnabled = False
        chkPause.IsEnabled = False
        chkRestrictMaxScore.IsEnabled = False
        chkRestrictMinScore.IsEnabled = False
        chkRestrictPageCount.IsEnabled = False
        chkRestrictMinWilsonScore.IsEnabled = False
        chkUseTrixieBooru.IsEnabled = False
        chkFilenameNoTags.IsEnabled = False
        chkSaveMetadataToFile.IsEnabled = False
        chkThumbnailOnly.IsEnabled = False
        chkResizeThumbnail.IsEnabled = False
        txtThumbnailWidth.IsEnabled = False
        txtThumbnailHeight.IsEnabled = False
        chkKeepThumbnailAspectRatio.IsEnabled = False
        cmbThumbnailFillColor.IsEnabled = False
        chkCacheAllPages.IsEnabled = False
        cmbFilters.IsEnabled = False
        cmbSordField.IsEnabled = False
        cmbSortDirection.IsEnabled = False
        cmbTagSeparator.IsEnabled = False
        sldMinWilsonScore.IsEnabled = False
    End Sub
    Private Sub UnlockWindow()
        txtMaxScore.IsEnabled = chkRestrictMaxScore.IsChecked
        txtMinScore.IsEnabled = chkRestrictMinScore.IsChecked
        txtPauseDuration.IsEnabled = chkPause.IsChecked
        txtPauseThreshold.IsEnabled = chkPause.IsChecked
        txtSaveTo.IsEnabled = True
        txtSearchKey.IsEnabled = True
        txtPageIndexBegin.IsEnabled = chkRestrictPageCount.IsChecked
        txtPageIndexEnd.IsEnabled = chkRestrictPageCount.IsChecked
        btnBrowse.IsEnabled = True
        btnStart.IsEnabled = True
        chkPause.IsEnabled = True
        chkRestrictMaxScore.IsEnabled = True
        chkRestrictMinScore.IsEnabled = True
        chkRestrictPageCount.IsEnabled = True
        chkRestrictMinWilsonScore.IsEnabled = True
        chkUseTrixieBooru.IsEnabled = True
        chkFilenameNoTags.IsEnabled = True
        chkSaveMetadataToFile.IsEnabled = True
        chkThumbnailOnly.IsEnabled = True
        chkResizeThumbnail.IsEnabled = chkThumbnailOnly.IsChecked
        txtThumbnailWidth.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        txtThumbnailHeight.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        chkKeepThumbnailAspectRatio.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        cmbThumbnailFillColor.IsEnabled = chkThumbnailOnly.IsChecked And chkKeepThumbnailAspectRatio.IsChecked
        chkCacheAllPages.IsEnabled = True
        cmbFilters.IsEnabled = True
        cmbSordField.IsEnabled = True
        cmbSortDirection.IsEnabled = True
        cmbTagSeparator.IsEnabled = chkSaveMetadataToFile.IsChecked
        sldMinWilsonScore.IsEnabled = chkRestrictMinWilsonScore.IsChecked
    End Sub
    Private Sub SaveSettings()
        SaveSetting(ApplicationName, SettingsSectionName, LastDownloadPathKey, txtSaveTo.Text)
    End Sub
    Private Sub LoadSettings()
        DownloadedFileSavePath = GetSetting(ApplicationName, SettingsSectionName, LastDownloadPathKey, LastDownloadPathDefVal)
        txtSaveTo.Text = DownloadedFileSavePath
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
            Try
                .SelectedPath = txtSaveTo.Text
            Catch ex As Exception

            End Try
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            DownloadedFileSavePath = FolderBrowser.SelectedPath
            If DownloadedFileSavePath(DownloadedFileSavePath.Length - 1) <> "\" Then
                DownloadedFileSavePath = DownloadedFileSavePath & "\"
            End If
            txtSaveTo.Text = DownloadedFileSavePath
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As RoutedEventArgs) Handles btnStart.Click
        LockWindow()
        SaveSettings()
        prgProgress.Minimum = 0
        prgProgress.Maximum = 100
        prgProgress.Value = 0
        '檢查搜尋條件是否有效
        If txtSearchKey.Text.Trim() = "" Then
            MessageBox.Show("請輸入搜尋條件。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockWindow()
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End If
        '檢查擱置時間是否有效
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
            PauseThreshold = Math.Abs(Int(txtPauseThreshold.Text))
            PauseDuration = Math.Abs(Int(txtPauseDuration.Text))
        End If
        '檢查下載頁數是否有效
        If chkRestrictPageCount.IsChecked Then
            If Not IsNumeric(txtPageIndexBegin.Text) Or Not IsNumeric(txtPageIndexEnd.Text) Then
                MessageBox.Show("指定的下載頁數存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Int(txtPageIndexBegin.Text) < 1 Or Int(txtPageIndexEnd.Text) < 1 Then
                MessageBox.Show("指定的下載頁數必須為正數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            If Int(txtPageIndexBegin.Text) > Int(txtPageIndexEnd.Text) Then
                MessageBox.Show("開始的下載頁數必須不大於結束的頁數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            PageIndexBegin = Int(txtPageIndexBegin.Text)
            PageIndexEnd = Int(txtPageIndexEnd.Text)
            PageCount = PageIndexEnd - PageIndexBegin + 1
        End If
        '構造檢索請求URL
        Dim iPageIndex As Integer = 1
        Dim Filter As New ComboBoxItem
        Dim SortField As ComboBoxItem = cmbSordField.SelectedItem
        Dim SortDirection As ComboBoxItem = cmbSortDirection.SelectedItem
        Filter = cmbFilters.SelectedItem
        SearchQuery = CurrentSearchPrefix & txtSearchKey.Text.Trim().Replace(" ", "+") & _
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
            MinScore = txtMinScore.Text
            'sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString()
        End If
        If chkRestrictMaxScore.IsChecked Then
            If Not IsNumeric(txtMaxScore.Text) Then
                MessageBox.Show("指定的最高評分值存在錯誤。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockWindow()
                SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                Exit Sub
            End If
            MaxScore = txtMaxScore.Text
            'sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString()
        End If
        If chkRestrictMaxScore.IsChecked And chkRestrictMinScore.IsChecked Then
            If MinScore > MaxScore Then
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
        '執行第一次檢索，擷取圖像總數資訊
        URLList.Clear()
        RefreshURLList()
        JSONCache.Clear()
        txtStatus.Text = "正在連線到 " & SearchQuery & " 以 GET 方法擷取 JSON 檔案。"
        UpdateLayout()
        Dim sJSON As String = ""
        Dim iPageTotal As Integer
        Dim HttpReq As New SimpleHttpRequest.HttpJsonRequest
        Try
            sJSON = HttpReq.DoGet(SearchQuery)
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
        Dim iImageCountOnLastPage As Integer = IIf(iImageTotal Mod 50 = 0, 50, iImageTotal Mod 50)
        iPageTotal = Math.Ceiling((CInt(JSONResponse("total")) / 50))
        Dim iTotalPageCountToDownload As Integer = iPageTotal
        Dim iTotalImageCountToDownload As Integer = iImageTotal
        txtStatus.Text = "JSON 檔案擷取完畢，總共搜尋到 " & JSONResponse("total").ToString & " 張相片。一共 " & iPageTotal.ToString & " 個分頁。"
        UpdateLayout()
        If JSONResponse("total").ToString = "0" Then
            MessageBox.Show("沒有找到符合您指定的搜尋條件的相片。", "沒有結果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            UnlockWindow()
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            Exit Sub
        End If
        If MessageBox.Show("搜尋完畢!" & vbCrLf & "本次搜尋總共搜尋到 " & JSONResponse("total").ToString & " 張相片。一共 " & iPageTotal.ToString & " 個分頁。" & vbCrLf & "開始下載相片嗎?", "下載相片", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Forms.DialogResult.Yes Then
            '快取搜尋結果
            If chkCacheAllPages.IsChecked Then
                txtStatus.Text = "正在快取搜尋結果。"
                UpdateLayout()
                Try
                    sJSON = HttpReq.DoGet(SearchQuery)
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
                iTotalPageCountToDownload = iPageTotal
                iTotalImageCountToDownload = iImageTotal
                '計算實際上需要下載的頁面編號
                If chkRestrictPageCount.IsChecked Then
                    If PageIndexEnd < iPageTotal Then
                        iTotalPageCountToDownload = PageCount
                        iTotalImageCountToDownload = PageCount * 50
                    ElseIf PageIndexEnd <= iPageTotal Then
                        iTotalPageCountToDownload = PageCount
                        iTotalImageCountToDownload = (PageCount - 1) * 50 + iImageCountOnLastPage
                    ElseIf PageIndexBegin <= iPageTotal And PageIndexEnd > iPageTotal Then
                        PageIndexEnd = iPageTotal
                        PageCount = PageIndexEnd - PageIndexBegin + 1
                        iTotalPageCountToDownload = PageCount
                        iTotalImageCountToDownload = (PageCount - 1) * 50 + iImageCountOnLastPage
                    Else
                        MessageBox.Show("開始的下載頁數必須不大於搜尋結果的最大頁數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtStatus.Text = "開始的下載頁數必須不大於搜尋結果的最大頁數。"
                        UnlockWindow()
                        SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                        Exit Sub
                    End If
                End If
                prgProgress.Maximum = iTotalPageCountToDownload
                prgProgress.Minimum = 0
                prgProgress.Value = 0
                For iPageIndex = PageIndexBegin To PageIndexEnd
                    SearchQuery = CurrentSearchPrefix & txtSearchKey.Text.Trim().Replace(" ", "+") & _
                                   DerpibooruSearchPageSelector & iPageIndex.ToString() & _
                                   DerpibooruImagesPerpageSelector & _
                                   DerpibooruImagesSortFieldSelector & SortField.Tag.ToString & _
                                   DerpibooruImagesSortDirectionSelector & SortDirection.Tag.ToString & _
                                   DerpibooruImagesFilterSelector & IIf(Filter.Tag = -1, UserSpecifiedFilterID.ToString, Filter.Tag.ToString)
                    '由於API變更,已廢棄
                    'If chkRestrictMinScore.IsChecked Then
                    '    iMinScore = txtMinScore.Text
                    '    'sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString()
                    'End If
                    'If chkRestrictMaxScore.IsChecked Then
                    '    iMaxScore = txtMaxScore.Text
                    '    'sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString()
                    'End If
                    Try
                        sJSON = HttpReq.DoGet(SearchQuery)
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
                    prgProgress.Value = iPageIndex - PageIndexBegin + 1
                    SetTaskbarProgess(iTotalPageCountToDownload, 0, iPageIndex - PageIndexBegin + 1, Shell.TaskbarItemProgressState.Paused)
                Next
            End If
            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
            txtStatus.Text = "正在從 Derpibooru 下載檔案。"
            If Not chkCacheAllPages.IsChecked Then
                '未使用快取時，重新計算資料
                Try
                    sJSON = HttpReq.DoGet(SearchQuery)
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
                iTotalPageCountToDownload = iPageTotal
                iTotalImageCountToDownload = iImageTotal
                '計算實際上需要下載的頁面編號
                If chkRestrictPageCount.IsChecked Then
                    If PageIndexEnd < iPageTotal Then
                        iTotalPageCountToDownload = PageCount
                        iTotalImageCountToDownload = PageCount * 50
                    ElseIf PageIndexEnd <= iPageTotal Then
                        iTotalPageCountToDownload = PageCount
                        iTotalImageCountToDownload = (PageCount - 1) * 50 + iImageCountOnLastPage
                    ElseIf PageIndexBegin <= iPageTotal And PageIndexEnd > iPageTotal Then
                        PageIndexEnd = iPageTotal
                        PageCount = PageIndexEnd - PageIndexBegin + 1
                        iTotalPageCountToDownload = PageCount
                        iTotalImageCountToDownload = (PageCount - 1) * 50 + iImageCountOnLastPage
                    Else
                        MessageBox.Show("開始的下載頁數必須不大於搜尋結果的最大頁數。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UnlockWindow()
                        SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                        Exit Sub
                    End If
                End If
            End If
            prgProgress.Maximum = iTotalImageCountToDownload
            prgProgress.Minimum = 0
            prgProgress.Value = 0
            Dim sImageFileName As String = ""
            Dim sImageURL As String = ""
            Dim nSuccess As Integer = 0
            Dim nFail As Integer = 0
            Dim nIgnored As Integer = 0
            Dim iPageIndexLoopBegin As Integer
            Dim iPageIndexLoopEnd As Integer
            If chkCacheAllPages.IsChecked Then
                iPageIndexLoopBegin = 1
                iPageIndexLoopEnd = iTotalPageCountToDownload
            Else
                iPageIndexLoopBegin = PageIndexBegin
                iPageIndexLoopEnd = PageIndexEnd
            End If
            '周遊每一個頁面
            For iPageIndex = iPageIndexLoopBegin To iPageIndexLoopEnd
                '決定是使用快取或是重新擷取資料
                If chkCacheAllPages.IsChecked Then
                    sJSON = JSONCache(iPageIndex - 1)
                Else
                    SearchQuery = CurrentSearchPrefix & txtSearchKey.Text.Trim().Replace(" ", "+") & _
                                   DerpibooruSearchPageSelector & iPageIndex.ToString() & _
                                   DerpibooruImagesPerpageSelector & _
                                   DerpibooruImagesSortFieldSelector & SortField.Tag.ToString & _
                                   DerpibooruImagesSortDirectionSelector & SortDirection.Tag.ToString & _
                                   DerpibooruImagesFilterSelector & IIf(Filter.Tag = -1, UserSpecifiedFilterID.ToString, Filter.Tag.ToString)
                    '由於API變更,已廢棄
                    'If chkRestrictMinScore.IsChecked Then
                    '    iMinScore = txtMinScore.Text
                    '    'sSearchQuery = sSearchQuery & DerpibooruImagesMinScoreSelector & iMinScore.ToString()
                    'End If
                    'If chkRestrictMaxScore.IsChecked Then
                    '    iMaxScore = txtMaxScore.Text
                    '    'sSearchQuery = sSearchQuery & DerpibooruImagesMaxScoreSelector & iMaxScore.ToString()
                    'End If
                    Try
                        sJSON = HttpReq.DoGet(SearchQuery)
                    Catch ex As Exception
                        Dim iIgnoredImageCount As Integer
                        If iPageIndex = iPageTotal Then
                            iIgnoredImageCount = iImageCountOnLastPage
                        Else
                            iIgnoredImageCount = 50
                        End If
                        nIgnored += iIgnoredImageCount
                        URLList.Add("嘗試擷取頁面" & iPageIndex.ToString & "的資訊時發生例外情況:" & vbCrLf & ex.Message & "，已略過 " & iIgnoredImageCount & " 個下載作業。")
                        RefreshURLList()
                        System.Windows.Forms.Application.DoEvents()
                        prgProgress.Value += iIgnoredImageCount
                        SetTaskbarProgess(iTotalImageCountToDownload, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                        UpdateLayout()
                        Continue For
                    End Try
                End If
                '讀取伺服器回應的資料
                JSONResponse = JsonConvert.DeserializeObject(sJSON)
                '周遊每一張影像
                For Each ImageJSON As JToken In JSONResponse("images").ToArray
                    '建立下載資料夾
                    If Not Directory.Exists(DownloadedFileSavePath) Then
                        Try
                            Directory.CreateDirectory(DownloadedFileSavePath)
                        Catch ex As Exception
                            MessageBox.Show("無法建立下載資料夾 '" & DownloadedFileSavePath & "'，發生例外情況: " & ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            URLList.Add("無法建立下載資料夾 '" & DownloadedFileSavePath & "'，發生例外情況: " & ex.Message)
                            RefreshURLList()
                            UnlockWindow()
                            SetTaskbarProgess(100, 0, 0, Shell.TaskbarItemProgressState.None)
                            Exit Sub
                        End Try
                    End If
                    '獲取下載URL與目標
                    If chkFilenameNoTags.IsChecked And Not chkThumbnailOnly.IsChecked Then
                        sImageURL = ImageJSON("representations")("full").ToString
                    ElseIf chkFilenameNoTags.IsChecked And chkThumbnailOnly.IsChecked Then
                        sImageURL = ImageJSON("representations")("medium").ToString
                    ElseIf Not chkFilenameNoTags.IsChecked And chkThumbnailOnly.IsChecked Then
                        sImageURL = ImageJSON("representations")("medium").ToString
                    Else
                        sImageURL = ImageJSON("view_url").ToString
                    End If
                    If chkFilenameNoTags.IsChecked Then
                        sImageFileName = GetFileNameFromDircectURL(ImageJSON("representations")("full").ToString)
                    Else
                        sImageFileName = GetFileNameFromDircectURL(ImageJSON("view_url").ToString)
                    End If
                    'MessageBox.Show(sImageFileName)
                    '檢查是否需要下載
                    If chkRestrictMinScore.IsChecked Then
                        If CInt(ImageJSON("score")) < MinScore Then
                            nIgnored += 1
                            URLList.Add("已略過 " & sImageURL & " 的下載作業。因為其評分 " & ImageJSON("score").ToString() & " 低於指定的最低評分 " & MinScore.ToString() & "。")
                            RefreshURLList()
                            System.Windows.Forms.Application.DoEvents()
                            prgProgress.Value += 1
                            SetTaskbarProgess(iTotalImageCountToDownload, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                            UpdateLayout()
                            If chkPause.IsChecked Then
                                If prgProgress.Value Mod PauseThreshold = 0 Then
                                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(PauseDuration))
                                End If
                            End If
                            Continue For
                        End If
                    End If
                    If chkRestrictMaxScore.IsChecked Then
                        If CInt(ImageJSON("score")) > MaxScore Then
                            nIgnored += 1
                            URLList.Add("已略過 " & sImageURL & " 的下載作業。因為其評分 " & ImageJSON("score").ToString() & " 高於指定的最高評分 " & MaxScore.ToString() & "。")
                            RefreshURLList()
                            System.Windows.Forms.Application.DoEvents()
                            prgProgress.Value += 1
                            SetTaskbarProgess(iTotalImageCountToDownload, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                            UpdateLayout()
                            If chkPause.IsChecked Then
                                If prgProgress.Value Mod PauseThreshold = 0 Then
                                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(PauseDuration))
                                End If
                            End If
                            Continue For
                        End If
                    End If
                    If chkRestrictMinWilsonScore.IsChecked Then
                        If CDbl(ImageJSON("wilson_score")) < sldMinWilsonScore.Value Then
                            nIgnored += 1
                            URLList.Add("已略過 " & sImageURL & " 的下載作業。因為其質量評分 " & ImageJSON("wilson_score").ToString() & " 低於指定的最低質量評分 " & sldMinWilsonScore.Value.ToString() & "。")
                            RefreshURLList()
                            System.Windows.Forms.Application.DoEvents()
                            prgProgress.Value += 1
                            SetTaskbarProgess(iTotalImageCountToDownload, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                            UpdateLayout()
                            If chkPause.IsChecked Then
                                If prgProgress.Value Mod PauseThreshold = 0 Then
                                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(PauseDuration))
                                End If
                            End If
                            Continue For
                        End If
                    End If
                    '下載影像
                    Dim FileDownloader As New WebClient
                    Try
                        '刪除已經存在的檔案
                        If File.Exists(DownloadedFileSavePath & sImageFileName) Then
                            Try
                                File.Delete(DownloadedFileSavePath & sImageFileName)
                            Catch ex As Exception

                            End Try
                        End If
                        '下載檔案
                        FileDownloader.DownloadFile(sImageURL, DownloadedFileSavePath & sImageFileName)
                        '一致化縮圖
                        If chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked Then
                            '讀取原始圖像
                            Dim SourceBitmap As Bitmap = New Bitmap(DownloadedFileSavePath & sImageFileName)

                            '決定一致化策略
                            Dim NormalizedBitmap As Bitmap
                            If chkKeepThumbnailAspectRatio.IsChecked Then
                                '若使用者不需要保持外觀比例，填充或裁切圖像
                                Dim PaddedBitmapWidth As Integer
                                Dim PaddedBitmapHeight As Integer
                                If SourceBitmap.Width >= SourceBitmap.Height Then
                                    PaddedBitmapWidth = SourceBitmap.Width
                                    PaddedBitmapHeight = SourceBitmap.Width * ThumbnailHeight / ThumbnailWidth
                                Else
                                    PaddedBitmapHeight = SourceBitmap.Height
                                    PaddedBitmapWidth = SourceBitmap.Height * ThumbnailWidth / ThumbnailHeight
                                End If
                                Dim ThumbnailFillColorItem As ComboBoxItem = cmbThumbnailFillColor.SelectedItem
                                Dim ThumbnailFillColor As Color = ThumbnailFillColorItem.Tag
                                Dim PaddedBitmap As Bitmap = PadBitmap(SourceBitmap, PaddedBitmapWidth, PaddedBitmapHeight, ThumbnailFillColor, ContentAlignment.MiddleCenter)
                                NormalizedBitmap = RescaleBitmap(PaddedBitmap, ThumbnailWidth, ThumbnailHeight)
                            Else
                                '若使用者不需要保持外觀比例，直接調整圖像尺寸
                                NormalizedBitmap = RescaleBitmap(SourceBitmap, ThumbnailWidth, ThumbnailHeight)
                            End If

                            '儲存到暫存檔
                            Dim NormalizedBitmapEncoderParams As New EncoderParameters(1)
                            NormalizedBitmapEncoderParams.Param(0) = New EncoderParameter(Encoder.Quality, Int(100))
                            If File.Exists(DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & "_tmp.png") Then
                                Try
                                    File.Delete(DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & "_tmp.png")
                                Catch ex As Exception

                                End Try
                            End If
                            NormalizedBitmap.Save(DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & "_tmp.png", GetImageEncoderInfo(ImageFormat.Png), NormalizedBitmapEncoderParams)
                            '取代下載的文件
                            SourceBitmap.Dispose()
                            Try
                                File.Delete(DownloadedFileSavePath & sImageFileName)
                            Catch ex As Exception

                            End Try
                            Try
                                File.Delete(DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & ".png")
                            Catch ex As Exception

                            End Try
                            File.Move(DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & "_tmp.png", DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & ".png")
                            Try
                                File.Delete(DownloadedFileSavePath & Path.GetFileNameWithoutExtension(sImageFileName) & "_tmp.png")
                            Catch ex As Exception

                            End Try
                            sImageFileName = Path.GetFileNameWithoutExtension(sImageFileName) & ".png"
                        End If
                        '儲存標籤
                        If chkSaveMetadataToFile.IsChecked Then
                            Dim MetadataFilePath As String = DownloadedFileSavePath & IO.Path.GetFileNameWithoutExtension(sImageFileName) & ".txt"
                            Dim MetadataFileStream As New FileStream(MetadataFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                            Dim MetadataFileWriter As New StreamWriter(MetadataFileStream)
                            Dim TagList As List(Of JToken) = ImageJSON("tags").ToList
                            If TagList.Count > 0 Then
                                For i As Integer = 0 To TagList.Count - 1
                                    '寫標籤
                                    Select Case cmbTagSeparator.SelectedIndex
                                        Case 0 '逗號
                                            MetadataFileWriter.Write(TagList(i).ToString())
                                        Case 1 '換行
                                            MetadataFileWriter.Write(TagList(i).ToString())
                                        Case 2 '空白
                                            MetadataFileWriter.Write(TagList(i).ToString().Replace(" ", "_"))
                                        Case Else
                                            MetadataFileWriter.Write(TagList(i))
                                    End Select
                                    '寫分隔符號
                                    If i <> TagList.Count - 1 Then
                                        Select Case cmbTagSeparator.SelectedIndex
                                            Case 0 '逗號
                                                MetadataFileWriter.Write(", ")
                                            Case 1 '換行
                                                MetadataFileWriter.Write(vbCrLf)
                                            Case 2 '空白
                                                MetadataFileWriter.Write(" ")
                                            Case Else
                                                MetadataFileWriter.Write(", ")
                                        End Select
                                    End If
                                Next
                            End If
                            MetadataFileWriter.Close()
                            MetadataFileStream.Close()
                        End If
                        nSuccess += 1
                        URLList.Add("成功從 " & sImageURL & " 下載相片到 " & DownloadedFileSavePath & sImageFileName)
                        RefreshURLList()
                    Catch ex As Exception
                        URLList.Add("從 " & sImageURL & " 下載相片時失敗，發生例外情況: " & ex.Message)
                        RefreshURLList()
                        nFail += 1
                    End Try
                    System.Windows.Forms.Application.DoEvents()
                    prgProgress.Value += 1
                    SetTaskbarProgess(iTotalImageCountToDownload, 0, prgProgress.Value, Shell.TaskbarItemProgressState.Normal)
                    UpdateLayout()
                    If chkPause.IsChecked Then
                        If prgProgress.Value Mod PauseThreshold = 0 Then
                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(PauseDuration))
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

    Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        SaveSettings()
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurituProctocolTypeExtensions.Tls11 Or SecurituProctocolTypeExtensions.Tls12
        Catch ex As Exception
            MessageBox.Show("無法配置應用程式以啟用對 TLS 1.1 和 TLS 1.2 的支援，因為發生例外情況:" & vbCrLf & ex.Message & vbCrLf & vbCrLf & "如果您正在使用 Windows 7 或更早版本的 Windows 作業系統，那麼您可能需要更新您的作業系統。" & vbCrLf & "應用程式仍將繼續啟動，但是可能無法正常使用。", "無法啟用必要的協力組件", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        LoadSettings()
        sldMinWilsonScore.IsEnabled = chkRestrictMinWilsonScore.IsChecked
        If chkUseTrixieBooru.IsChecked Then
            CurrentSearchPrefix = DerpibooruSearchPrefixBackup
        Else
            CurrentSearchPrefix = DerpibooruSearchPrefix
        End If
        cmbTagSeparator.IsEnabled = chkSaveMetadataToFile.IsChecked
        txtSearchKey.Focus()
    End Sub

    Private Sub chkRestrictPageCount_Click(sender As Object, e As RoutedEventArgs) Handles chkRestrictPageCount.Click
        txtPageIndexBegin.IsEnabled = chkRestrictPageCount.IsChecked
        txtPageIndexEnd.IsEnabled = chkRestrictPageCount.IsChecked
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

    Private Sub chkSaveMetadataToFile_Click(sender As Object, e As RoutedEventArgs) Handles chkSaveMetadataToFile.Click
        cmbTagSeparator.IsEnabled = chkSaveMetadataToFile.IsChecked
    End Sub

    Private Sub chkThumbnailOnly_Click(sender As Object, e As RoutedEventArgs) Handles chkThumbnailOnly.Click
        chkResizeThumbnail.IsEnabled = chkThumbnailOnly.IsChecked
        txtThumbnailWidth.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        txtThumbnailHeight.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        chkKeepThumbnailAspectRatio.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        cmbThumbnailFillColor.IsEnabled = chkThumbnailOnly.IsChecked And chkKeepThumbnailAspectRatio.IsChecked
    End Sub

    Private Sub chkResizeThumbnail_Click(sender As Object, e As RoutedEventArgs) Handles chkResizeThumbnail.Click
        chkResizeThumbnail.IsEnabled = chkThumbnailOnly.IsChecked
        txtThumbnailWidth.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        txtThumbnailHeight.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        chkKeepThumbnailAspectRatio.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        cmbThumbnailFillColor.IsEnabled = chkThumbnailOnly.IsChecked And chkKeepThumbnailAspectRatio.IsChecked
    End Sub

    Private Sub chkKeepThumbnailAspectRatio_Click(sender As Object, e As RoutedEventArgs) Handles chkKeepThumbnailAspectRatio.Click
        chkResizeThumbnail.IsEnabled = chkThumbnailOnly.IsChecked
        txtThumbnailWidth.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        txtThumbnailHeight.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        chkKeepThumbnailAspectRatio.IsEnabled = chkThumbnailOnly.IsChecked And chkResizeThumbnail.IsChecked
        cmbThumbnailFillColor.IsEnabled = chkThumbnailOnly.IsChecked And chkKeepThumbnailAspectRatio.IsChecked
    End Sub

    Private Sub txtThumbnailWidth_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtThumbnailWidth.LostFocus
        Try
            ThumbnailWidth = Int(txtThumbnailWidth.Text)
            If ThumbnailWidth <= 0 Then
                ThumbnailWidth = 1
            End If
        Catch ex As Exception
            ThumbnailWidth = 1
        End Try
        txtThumbnailWidth.Text = ThumbnailWidth.ToString()
    End Sub

    Private Sub txtThumbnailHeight_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtThumbnailHeight.LostFocus
        Try
            ThumbnailHeight = Int(txtThumbnailHeight.Text)
            If ThumbnailHeight <= 0 Then
                ThumbnailHeight = 1
            End If
        Catch ex As Exception
            ThumbnailHeight = 1
        End Try
        txtThumbnailHeight.Text = ThumbnailHeight.ToString()
    End Sub
End Class
