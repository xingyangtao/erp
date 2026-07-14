' ============================================================================
' 主窗口
' 功能：导航树、首页仪表盘、数据表格、分页
' 样式参考囍铺黄金 Web 仪表盘，数据逻辑移植自易语言主窗口
' ============================================================================

Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms.DataVisualization.Charting

Public Class MainForm
    Inherits System.Windows.Forms.Form

    Private Shared ReadOnly BgColor As Color = Color.FromArgb(240, 242, 245)
    Private Shared ReadOnly CardColor As Color = Color.White
    Private Shared ReadOnly PrimaryBlue As Color = Color.FromArgb(24, 144, 255)
    Private Shared ReadOnly TextDark As Color = Color.FromArgb(51, 51, 51)
    Private Shared ReadOnly TextMuted As Color = Color.FromArgb(140, 140, 140)
    Private Shared ReadOnly TrendReceivedColor As Color = Color.FromArgb(30, 144, 255)
    Private Shared ReadOnly TrendSalesColor As Color = Color.FromArgb(255, 105, 180)
    Private Shared ReadOnly TrendRecoveryColor As Color = Color.FromArgb(0, 139, 139)

    ' ========== 共享属性，供子窗口获取操作模式 ==========
    Public Shared ReturnOrderButtonName As String = ""       ' 退库操作按钮名称（添加/详情/再次提交）

    Private currentTableName As String = ""
    Private currentSortField As String = "id"
    Private currentSortOrder As String = "ASC"
    Private currentPage As Integer = 1
    Private pageSize As Integer = 50
    Private totalRecords As Integer = 0
    Private totalPages As Integer = 0
    Private currentPageName As String = ""
    Private currentPageID As String = ""

    Private treeView As New TreeView()
    Friend dgvMain As New DataGridView()         ' 主数据表格（Friend供子窗口访问）
    Private dgvRanking As DataGridView
    Private lblTitle As New Label()
    Private lblUserInfo As New Label()
    Private lblPostInfo As New Label()
    Private lblWelcome As New Label()
    Private lblPageInfo As New Label()
    Private WithEvents btnFirst As New Button()
    Private WithEvents btnPrev As New Button()
    Private WithEvents btnNext As New Button()
    Private WithEvents btnLast As New Button()
    Private cmbPageSize As New ComboBox()

    Private homePanel As Panel
    Private panelRight As Panel
    Private panelContent As Panel
    Private panelPage As Panel
    Private isUiInitializing As Boolean = False
    Private lblHomeWelcome As Label
    Private lblKpiValues As Label()
    Private chartTrend As Chart
    Private chartPie As Chart
    Private dgvRanking As DataGridView
    Private panelPieLegend As FlowLayoutPanel
    Private chkHideValues As Panel
    Private chkAutoPrint As Panel
    Private hideValuesEnabled As Boolean = False
    Private autoPrintEnabled As Boolean = False
    Private dashboardTimer As Timer
    Private trendPoints As List(Of HomeDashboardService.TrendPoint)
    Private isDashboardLoading As Boolean = False

    Public Sub New()
        InitializeUI()
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf MainForm_Resize
        AddHandler Me.FormClosing, AddressOf MainForm_FormClosing
        AddHandler treeView.AfterSelect, AddressOf TreeView_AfterSelect
        AddHandler Application.ThreadException, AddressOf App_ThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf App_UnhandledException
    End Sub

    ' ========== 子文件夹表格刷新（供子窗口调用） ==========
    Public Sub RefreshSubFolderTable()
        If dgvMain.Visible AndAlso dgvMain.Rows.Count > 0 Then
            Dim saveRow As Integer = If(dgvMain.CurrentCell?.RowIndex, 0)
            Form_Load(Nothing, Nothing)
            If saveRow < dgvMain.Rows.Count Then
                dgvMain.CurrentCell = dgvMain.Rows(saveRow).Cells(0)
            End If
        End If
    End Sub

    Private Sub App_ThreadException(sender As Object, e As System.Threading.ThreadExceptionEventArgs)
        Dim errorMsg As String = $"[{DateTime.Now}] ThreadException: {e.Exception.Message}" & vbCrLf & e.Exception.StackTrace & vbCrLf
        IO.File.AppendAllText(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error.log"), errorMsg)
        MessageBox.Show("系统错误: " & e.Exception.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub App_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Dim ex As Exception = TryCast(e.ExceptionObject, Exception)
        If ex IsNot Nothing Then
            Dim errorMsg As String = $"[{DateTime.Now}] UnhandledException: {ex.Message}" & vbCrLf & ex.StackTrace & vbCrLf
            IO.File.AppendAllText(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error.log"), errorMsg)
        End If
    End Sub

    Private Sub InitializeUI()
        Me.Text = $"囍铺黄金 - {ERPCompanyName}"
        Me.Size = New Size(1400, 900)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.WindowState = FormWindowState.Maximized
        Me.BackColor = BgColor
        Me.Font = New Font("微软雅黑", 9)

        ' ========== 左右分割容器（先添加，后停靠） ==========
        Dim splitContainer As New SplitContainer()
        splitContainer.Dock = DockStyle.Fill
        splitContainer.SplitterDistance = 210
        splitContainer.FixedPanel = FixedPanel.Panel1
        splitContainer.BackColor = BgColor
        splitContainer.SplitterWidth = 4

        ' 左侧导航树
        treeView.Dock = DockStyle.Fill
        treeView.Font = New Font("微软雅黑", 10)
        treeView.BackColor = CardColor
        treeView.BorderStyle = BorderStyle.None
        treeView.ShowLines = False
        treeView.FullRowSelect = True
        treeView.HideSelection = False
        treeView.ItemHeight = 34
        splitContainer.Panel1.Padding = New Padding(0, 4, 0, 4)
        splitContainer.Panel1.BackColor = CardColor
        splitContainer.Panel1.Controls.Add(treeView)

        ' 右侧面板
        panelRight = New Panel()
        panelRight.Dock = DockStyle.Fill
        panelRight.BackColor = BgColor
        splitContainer.Panel2.Controls.Add(panelRight)

        ' 分页面板
        panelPage = New Panel()
        panelPage.Dock = DockStyle.Bottom
        panelPage.Height = 42
        panelPage.BackColor = CardColor
        panelPage.Padding = New Padding(12, 6, 12, 6)
        panelRight.Controls.Add(panelPage)

        ' 内容面板
        panelContent = New Panel()
        panelContent.Dock = DockStyle.Fill
        panelContent.BackColor = BgColor
        panelRight.Controls.Add(panelContent)

        ' 分页控件
        lblPageInfo.Text = "共 0 条记录"
        lblPageInfo.AutoSize = True
        lblPageInfo.Location = New Point(12, 12)
        lblPageInfo.ForeColor = TextMuted
        panelPage.Controls.Add(lblPageInfo)

        ConfigurePageButton(btnFirst, "首页", 180)
        ConfigurePageButton(btnPrev, "上一页", 250)
        ConfigurePageButton(btnNext, "下一页", 330)
        ConfigurePageButton(btnLast, "尾页", 410)
        panelPage.Controls.Add(btnFirst)
        panelPage.Controls.Add(btnPrev)
        panelPage.Controls.Add(btnNext)
        panelPage.Controls.Add(btnLast)

        cmbPageSize.Location = New Point(490, 9)
        cmbPageSize.Size = New Size(80, 25)
        cmbPageSize.DropDownStyle = ComboBoxStyle.DropDownList
        cmbPageSize.Items.AddRange({"25", "50", "100", "500", "1000", "5000"})
        cmbPageSize.SelectedItem = "50"
        AddHandler cmbPageSize.SelectedIndexChanged, Sub()
                                                         pageSize = Integer.Parse(cmbPageSize.SelectedItem.ToString())
                                                         currentPage = 1
                                                         LoadData()
                                                     End Sub
        panelPage.Controls.Add(cmbPageSize)

        ' 标题
        lblTitle.Text = "首页"
        lblTitle.Font = New Font("微软雅黑", 13, FontStyle.Bold)
        lblTitle.Dock = DockStyle.Top
        lblTitle.Height = 0
        lblTitle.Visible = False
        panelContent.Controls.Add(lblTitle)

        ' 数据表格
        dgvMain.Dock = DockStyle.Fill
        dgvMain.ReadOnly = True
        dgvMain.AllowUserToAddRows = False
        dgvMain.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvMain.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        dgvMain.BackgroundColor = CardColor
        dgvMain.BorderStyle = BorderStyle.None
        dgvMain.RowHeadersVisible = False
        dgvMain.EnableHeadersVisualStyles = False
        dgvMain.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)
        dgvMain.ColumnHeadersDefaultCellStyle.ForeColor = TextDark
        dgvMain.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)
        dgvMain.Visible = False
        panelContent.Controls.Add(dgvMain)

        ' 首页仪表盘
        BuildHomeDashboard()

        ' ========== 顶部面板（最后添加，停靠在顶部） ==========
        Dim panelTop As New Panel()
        panelTop.Dock = DockStyle.Top
        panelTop.Height = 56
        panelTop.BackColor = CardColor
        panelTop.Padding = New Padding(16, 8, 16, 8)

        ' 使用 TableLayoutPanel 布局顶部
        Dim topLayout As New TableLayoutPanel()
        topLayout.Dock = DockStyle.Fill
        topLayout.ColumnCount = 3
        topLayout.RowCount = 1
        topLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 300))  ' Logo + 公司名
        topLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))   ' 中间空白
        topLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 280))  ' 用户信息 + 退出
        topLayout.BackColor = Color.Transparent
        panelTop.Controls.Add(topLayout)

        ' 左侧: Logo + 公司名
        Dim panelLeft As New FlowLayoutPanel()
        panelLeft.FlowDirection = FlowDirection.LeftToRight
        panelLeft.Dock = DockStyle.Fill
        panelLeft.BackColor = Color.Transparent
        panelLeft.WrapContents = False
        panelLeft.AutoSize = False
        topLayout.Controls.Add(panelLeft, 0, 0)

        Dim picLogo As New PictureBox()
        picLogo.Size = New Size(36, 36)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BackColor = Color.Transparent
        picLogo.Margin = New Padding(0, 4, 8, 0)
        panelLeft.Controls.Add(picLogo)

        Dim lblCompany As New Label()
        lblCompany.Text = If(String.IsNullOrEmpty(ERPCompanyName), "囍铺黄金", ERPCompanyName)
        lblCompany.Font = New Font("微软雅黑", 14, FontStyle.Bold)
        lblCompany.ForeColor = PrimaryBlue
        lblCompany.AutoSize = True
        lblCompany.TextAlign = ContentAlignment.MiddleCenter
        lblCompany.Margin = New Padding(0, 6, 0, 0)
        panelLeft.Controls.Add(lblCompany)

        ' 中间空白
        Dim panelMiddle As New Panel()
        panelMiddle.Dock = DockStyle.Fill
        panelMiddle.BackColor = Color.Transparent
        topLayout.Controls.Add(panelMiddle, 1, 0)

        ' 右侧: 用户信息 + 退出按钮
        Dim panelRightTop As New FlowLayoutPanel()
        panelRightTop.FlowDirection = FlowDirection.RightToLeft
        panelRightTop.Dock = DockStyle.Fill
        panelRightTop.BackColor = Color.Transparent
        panelRightTop.WrapContents = False
        topLayout.Controls.Add(panelRightTop, 2, 0)

        Dim btnLogout As New Button()
        btnLogout.Text = "⏻"
        btnLogout.Font = New Font("Segoe UI Symbol", 12)
        btnLogout.FlatStyle = FlatStyle.Flat
        btnLogout.FlatAppearance.BorderSize = 0
        btnLogout.Size = New Size(36, 36)
        btnLogout.ForeColor = TextMuted
        btnLogout.Cursor = Cursors.Hand
        btnLogout.Margin = New Padding(8, 4, 0, 0)
        AddHandler btnLogout.Click, Sub() Me.Close()
        panelRightTop.Controls.Add(btnLogout)

        lblUserInfo.Text = $"用户名: {UserName} | {UserPostName}"
        lblUserInfo.Font = New Font("微软雅黑", 10)
        lblUserInfo.ForeColor = TextDark
        lblUserInfo.AutoSize = True
        lblUserInfo.TextAlign = ContentAlignment.MiddleRight
        lblUserInfo.Margin = New Padding(0, 8, 0, 0)
        panelRightTop.Controls.Add(lblUserInfo)

        ' 添加控件到窗体（顺序很重要：先Fill后Top）
        Me.Controls.Add(splitContainer)
        Me.Controls.Add(panelTop)
    End Sub

    Private Sub ConfigurePageButton(btn As Button, text As String, x As Integer)
        btn.Text = text
        btn.Location = New Point(x, 6)
        btn.Size = New Size(68, 28)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderColor = Color.FromArgb(217, 217, 217)
        btn.BackColor = CardColor
        btn.ForeColor = TextDark
    End Sub

    Private Sub BuildHomeDashboard()
        homePanel = New Panel()
        homePanel.Name = "homePanel"
        homePanel.Dock = DockStyle.Fill
        homePanel.BackColor = BgColor
        homePanel.AutoScroll = True
        homePanel.Visible = False
        panelContent.Controls.Add(homePanel)

        Dim headerPanel As New Panel()
        headerPanel.Location = New Point(16, 12)
        headerPanel.Size = New Size(1200, 36)
        headerPanel.BackColor = Color.Transparent
        headerPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        homePanel.Controls.Add(headerPanel)

        lblHomeWelcome = New Label()
        lblHomeWelcome.Text = $"欢迎回来，{UserName}{If(String.IsNullOrEmpty(UserPostName), "", UserPostName)}"
        lblHomeWelcome.Font = New Font("微软雅黑", 14, FontStyle.Bold)
        lblHomeWelcome.ForeColor = TextDark
        lblHomeWelcome.AutoSize = True
        lblHomeWelcome.Location = New Point(0, 4)
        headerPanel.Controls.Add(lblHomeWelcome)

        chkAutoPrint = CreateToggleSwitch("自动打印", False)
        chkAutoPrint.Location = New Point(860, 4)
        chkAutoPrint.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        headerPanel.Controls.Add(chkAutoPrint)

        chkHideValues = CreateToggleSwitch("数值隐藏", False)
        chkHideValues.Location = New Point(980, 4)
        chkHideValues.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        headerPanel.Controls.Add(chkHideValues)

        Dim kpiPanel As New FlowLayoutPanel()
        kpiPanel.Name = "kpiPanel"
        kpiPanel.BackColor = Color.Transparent
        kpiPanel.WrapContents = False
        kpiPanel.AutoScroll = True
        homePanel.Controls.Add(kpiPanel)

        Dim kpiDefs = {
            ("今日销售(元)", Color.FromArgb(64, 158, 255), "销"),
            ("今日销售(件)", Color.FromArgb(114, 46, 209), "件"),
            ("今日换购(g)", Color.FromArgb(250, 140, 22), "重"),
            ("今日回收(元)", Color.FromArgb(82, 196, 26), "回"),
            ("今日回收(g)", Color.FromArgb(19, 194, 194), "克"),
            ("实收金额(元)", Color.FromArgb(47, 84, 235), "收"),
            ("实际重量(g)", Color.FromArgb(235, 47, 150), "量")
        }

        ReDim lblKpiValues(kpiDefs.Length - 1)
        For i As Integer = 0 To kpiDefs.Length - 1
            Dim card = CreateKpiCard(kpiDefs(i).Item1, kpiDefs(i).Item2, kpiDefs(i).Item3)
            lblKpiValues(i) = DirectCast(card.Controls(2), Label)
            kpiPanel.Controls.Add(card)
        Next

        Dim quickPanel = CreateQuickAccessPanel()
        quickPanel.Location = New Point(1016, 56)
        quickPanel.Size = New Size(170, 520)
        quickPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        homePanel.Controls.Add(quickPanel)

        Dim chartCard = CreateCardPanel("实收数据(近30日)", New Point(16, 160), New Size(780, 300))
        chartCard.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        homePanel.Controls.Add(chartCard)

        chartTrend = CreateTrendChart()
        chartTrend.Location = New Point(12, 42)
        chartTrend.Size = New Size(chartCard.Width - 24, chartCard.Height - 54)
        chartTrend.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        chartCard.Controls.Add(chartTrend)

        Dim bottomLeft = CreateCardPanel("月销售排行榜", New Point(16, 470), New Size(500, 260))
        bottomLeft.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom
        homePanel.Controls.Add(bottomLeft)

        dgvRanking = New DataGridView()
        dgvRanking.Location = New Point(12, 42)
        dgvRanking.Size = New Size(bottomLeft.Width - 24, bottomLeft.Height - 54)
        dgvRanking.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        dgvRanking.ReadOnly = True
        dgvRanking.AllowUserToAddRows = False
        dgvRanking.RowHeadersVisible = False
        dgvRanking.BackgroundColor = CardColor
        dgvRanking.BorderStyle = BorderStyle.None
        dgvRanking.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvRanking.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvRanking.EnableHeadersVisualStyles = False
        dgvRanking.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)
        dgvRanking.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250)
        dgvRanking.Columns.Add("rank", "序号")
        dgvRanking.Columns.Add("item", "款号")
        dgvRanking.Columns.Add("category", "品类")
        dgvRanking.Columns.Add("spec", "规格")
        dgvRanking.Columns.Add("qty", "销量(件)")
        dgvRanking.Columns.Add("amount", "销售金额(元)")
        bottomLeft.Controls.Add(dgvRanking)

        Dim bottomRight = CreateCardPanel("当日销售类别占比(g)", New Point(530, 470), New Size(470, 260))
        bottomRight.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        homePanel.Controls.Add(bottomRight)

        chartPie = CreatePieChart()
        chartPie.Location = New Point(12, 42)
        chartPie.Size = New Size(280, bottomRight.Height - 54)
        chartPie.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom
        bottomRight.Controls.Add(chartPie)

        panelPieLegend = New FlowLayoutPanel()
        panelPieLegend.Location = New Point(300, 48)
        panelPieLegend.Size = New Size(150, bottomRight.Height - 60)
        panelPieLegend.Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        panelPieLegend.FlowDirection = FlowDirection.TopDown
        panelPieLegend.WrapContents = False
        panelPieLegend.AutoScroll = True
        panelPieLegend.BackColor = Color.Transparent
        bottomRight.Controls.Add(panelPieLegend)

        dashboardTimer = New Timer()
        dashboardTimer.Interval = 60000
        AddHandler dashboardTimer.Tick, AddressOf DashboardTimer_Tick
        LayoutHomeDashboard()
    End Sub

    Private Function CreateCardPanel(title As String, location As Point, size As Size) As Panel
        Dim card As New Panel()
        card.Location = location
        card.Size = size
        card.BackColor = CardColor
        card.Padding = New Padding(12)

        AddHandler card.Paint, Sub(s, e)
                                   Using pen As New Pen(Color.FromArgb(230, 230, 230))
                                       e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                                       e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1)
                                   End Using
                               End Sub

        Dim lbl As New Label()
        lbl.Text = title
        lbl.Font = New Font("微软雅黑", 11, FontStyle.Bold)
        lbl.ForeColor = TextDark
        lbl.AutoSize = True
        lbl.Location = New Point(12, 10)
        card.Controls.Add(lbl)
        Return card
    End Function

    Private Function CreateKpiCard(title As String, iconColor As Color, iconText As String) As Panel
        Dim card As New Panel()
        card.Size = New Size(148, 88)
        card.Margin = New Padding(0, 0, 10, 0)
        card.BackColor = CardColor
        card.Padding = New Padding(12, 10, 8, 8)

        AddHandler card.Paint, Sub(s, e)
                                   Using path As New GraphicsPath()
                                       Dim rect As New Rectangle(0, 0, card.Width - 1, card.Height - 1)
                                       path.AddRectangle(rect)
                                       Using pen As New Pen(Color.FromArgb(235, 235, 235))
                                           e.Graphics.DrawPath(pen, path)
                                       End Using
                                   End Using
                               End Sub

        Dim icon As New Panel()
        icon.Size = New Size(34, 34)
        icon.Location = New Point(12, 14)
        icon.BackColor = iconColor
        Dim iconLbl As New Label()
        iconLbl.Text = iconText
        iconLbl.ForeColor = Color.White
        iconLbl.Font = New Font("微软雅黑", 9, FontStyle.Bold)
        iconLbl.Dock = DockStyle.Fill
        iconLbl.TextAlign = ContentAlignment.MiddleCenter
        icon.Controls.Add(iconLbl)
        card.Controls.Add(icon)

        Dim titleLbl As New Label()
        titleLbl.Text = title
        titleLbl.ForeColor = TextMuted
        titleLbl.Font = New Font("微软雅黑", 8.5F)
        titleLbl.Location = New Point(54, 12)
        titleLbl.AutoSize = True
        card.Controls.Add(titleLbl)

        Dim valueLbl As New Label()
        valueLbl.Text = "0.00"
        valueLbl.ForeColor = TextDark
        valueLbl.Font = New Font("微软雅黑", 15, FontStyle.Bold)
        valueLbl.Location = New Point(54, 34)
        valueLbl.AutoSize = True
        card.Controls.Add(valueLbl)
        Return card
    End Function

    Private Function CreateToggleSwitch(caption As String, initialChecked As Boolean) As Panel
        Dim container As New Panel()
        container.Size = New Size(110, 28)
        container.BackColor = Color.Transparent
        container.Cursor = Cursors.Hand
        container.Tag = initialChecked

        Dim captionLbl As New Label()
        captionLbl.Text = caption
        captionLbl.ForeColor = TextMuted
        captionLbl.Font = New Font("微软雅黑", 9)
        captionLbl.AutoSize = True
        captionLbl.Location = New Point(0, 5)
        container.Controls.Add(captionLbl)

        Dim track As New Panel()
        track.Size = New Size(42, 22)
        track.Location = New Point(62, 3)
        track.BackColor = If(initialChecked, PrimaryBlue, Color.FromArgb(200, 200, 200))
        container.Controls.Add(track)

        Dim thumb As New Panel()
        thumb.Size = New Size(18, 18)
        thumb.Location = New Point(If(initialChecked, 22, 2), 2)
        thumb.BackColor = Color.White
        track.Controls.Add(thumb)

        AddHandler container.Click, Sub()
                                        Dim checked = Not CBool(container.Tag)
                                        container.Tag = checked
                                        track.BackColor = If(checked, PrimaryBlue, Color.FromArgb(200, 200, 200))
                                        thumb.Location = New Point(If(checked, 22, 2), 2)
                                        If container Is chkHideValues Then
                                            hideValuesEnabled = checked
                                            RefreshDashboardData()
                                        ElseIf container Is chkAutoPrint Then
                                            autoPrintEnabled = checked
                                        End If
                                    End Sub
        Return container
    End Function

    Private Function CreateQuickAccessPanel() As Panel
        Dim panel As New Panel()
        panel.BackColor = CardColor
        panel.Padding = New Padding(10)

        AddHandler panel.Paint, Sub(s, e)
                                    Using pen As New Pen(Color.FromArgb(230, 230, 230))
                                        e.Graphics.DrawRectangle(pen, 0, 0, panel.Width - 1, panel.Height - 1)
                                    End Using
                                End Sub

        Dim title As New Label()
        title.Text = "快捷入口"
        title.Font = New Font("微软雅黑", 11, FontStyle.Bold)
        title.ForeColor = TextDark
        title.Location = New Point(10, 10)
        title.AutoSize = True
        panel.Controls.Add(title)

        Dim actions = {
            ("个人信息", Sub() OpenChildForm(Of PersonalInfoForm)()),
            ("修改密码", Sub() OpenChildForm(Of PasswordChangeForm)()),
            ("款号合并", Sub() OpenChildForm(Of KuanHaoMergeForm)()),
            ("标签制作", Sub() OpenChildForm(Of BatchPrintForm)()),
            ("成本修改", Sub() OpenChildForm(Of ProductEditForm)()),
            ("报表打印", Sub() OpenChildForm(Of SalesReportForm)()),
            ("盘点清算", Sub() OpenChildForm(Of InventoryCheckForm)()),
            ("数据备份", Sub()
                             Using frm As New ExportDataForm(New DataTable())
                                 frm.ShowDialog(Me)
                             End Using
                         End Sub)
        }

        Dim y As Integer = 44
        For i As Integer = 0 To actions.Length - 1
            Dim row = i \ 2
            Dim col = i Mod 2
            Dim btn As New Panel()
            btn.Size = New Size(68, 72)
            btn.Location = New Point(10 + col * 78, 44 + row * 82)
            btn.BackColor = Color.FromArgb(248, 248, 248)
            btn.Cursor = Cursors.Hand
            btn.Tag = actions(i).Item2

            Dim icon As New Label()
            icon.Text = "⚙"
            icon.Font = New Font("Segoe UI Symbol", 14)
            icon.ForeColor = TextMuted
            icon.TextAlign = ContentAlignment.MiddleCenter
            icon.Dock = DockStyle.Top
            icon.Height = 36
            btn.Controls.Add(icon)

            Dim text As New Label()
            text.Text = actions(i).Item1
            text.Font = New Font("微软雅黑", 8.5F)
            text.ForeColor = TextDark
            text.TextAlign = ContentAlignment.TopCenter
            text.Dock = DockStyle.Fill
            btn.Controls.Add(text)

            AddHandler btn.Click, Sub(s, ev)
                                      Dim action = TryCast(DirectCast(s, Panel).Tag, Action)
                                      action?.Invoke()
                                  End Sub
            panel.Controls.Add(btn)
        Next
        Return panel
    End Function

    Private Sub OpenChildForm(Of T As {Form, New})()
        Using frm As New T()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Function CreateTrendChart() As Chart
        Dim chart As New Chart()
        chart.BackColor = CardColor
        chart.AntiAliasing = AntiAliasingStyles.All
        chart.TextAntiAliasingQuality = TextAntiAliasingQuality.High

        Dim area As New ChartArea("TrendArea")
        area.BackColor = CardColor
        area.AxisX.MajorGrid.LineColor = Color.FromArgb(240, 240, 240)
        area.AxisY.MajorGrid.LineColor = Color.FromArgb(240, 240, 240)
        area.AxisX.LabelStyle.ForeColor = TextMuted
        area.AxisY.LabelStyle.ForeColor = TextMuted
        area.AxisX.Interval = 5
        chart.ChartAreas.Add(area)

        AddLineSeries(chart, "实收", TrendReceivedColor)
        AddLineSeries(chart, "销售", TrendSalesColor)
        AddLineSeries(chart, "回收", TrendRecoveryColor)

        Dim legend As New Legend("Legend")
        legend.Docking = Docking.Right
        legend.Alignment = StringAlignment.Center
        legend.BackColor = Color.Transparent
        chart.Legends.Add(legend)
        Return chart
    End Function

    Private Sub AddLineSeries(chart As Chart, name As String, color As Color)
        Dim series As New Series(name)
        series.ChartType = SeriesChartType.Line
        series.ChartArea = "TrendArea"
        series.Color = color
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.Circle
        series.MarkerSize = 4
        chart.Series.Add(series)
    End Sub

    Private Function CreatePieChart() As Chart
        Dim chart As New Chart()
        chart.BackColor = CardColor
        Dim area As New ChartArea("PieArea")
        area.BackColor = CardColor
        chart.ChartAreas.Add(area)

        Dim series As New Series("Category")
        series.ChartType = SeriesChartType.Pie
        series.ChartArea = "PieArea"
        series("PieLabelStyle") = "Disabled"
        chart.Series.Add(series)
        chart.Legends.Clear()
        Return chart
    End Function

    Private Sub MainForm_Resize(sender As Object, e As EventArgs)
        If homePanel Is Nothing OrElse Not homePanel.Visible Then Return
        LayoutHomeDashboard()
    End Sub

    Private Sub LayoutHomeDashboard()
        Dim margin = 16
        Dim quickWidth = 170
        Dim contentWidth = Math.Max(homePanel.ClientSize.Width - margin * 2 - quickWidth - 12, 600)
        Dim rightX = margin + contentWidth + 12

        For Each ctrl As Control In homePanel.Controls
            If TypeOf ctrl Is FlowLayoutPanel AndAlso ctrl.Name = "kpiPanel" Then
                ctrl.Location = New Point(margin, 56)
                ctrl.Size = New Size(contentWidth, 96)
            End If
        Next

        For Each ctrl As Control In homePanel.Controls
            If TypeOf ctrl Is Panel AndAlso ctrl.BackColor = CardColor AndAlso ctrl.Controls.Count > 0 Then
                Dim title = TryCast(ctrl.Controls(0), Label)
                If title Is Nothing Then Continue For
                Select Case title.Text
                    Case "实收数据(近30日)"
                        ctrl.Location = New Point(margin, 160)
                        ctrl.Size = New Size(contentWidth, 300)
                    Case "月销售排行榜"
                        ctrl.Location = New Point(margin, 470)
                        ctrl.Size = New Size(CInt(contentWidth * 0.55), 260)
                    Case "当日销售类别占比(g)"
                        ctrl.Location = New Point(margin + CInt(contentWidth * 0.57), 470)
                        ctrl.Size = New Size(contentWidth - CInt(contentWidth * 0.57), 260)
                    Case "快捷入口"
                        ctrl.Location = New Point(rightX, 56)
                        ctrl.Size = New Size(quickWidth, Math.Max(homePanel.ClientSize.Height - 72, 420))
                End Select
            End If
        Next
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(GlobalViewSQL) Then BuildGlobalViewSQL()
        LoadNavigationTree()
        ShowHomePage()
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        If dashboardTimer IsNot Nothing Then
            dashboardTimer.Stop()
            dashboardTimer.Dispose()
        End If

        ' 关闭所有数据库连接
        CloseAllConnections()
    End Sub

    Private Sub LoadNavigationTree()
        Try
            isUiInitializing = True
            treeView.Nodes.Clear()

            Dim homeNode As New TreeNode("首页")
            homeNode.Tag = "0"
            treeView.Nodes.Add(homeNode)

            ' UserRole 格式: 'chakan1','chakan2','chakan3'（已带引号）
            Dim roleList As String = If(String.IsNullOrEmpty(UserRole), "''", UserRole)

            Dim sql = $"SELECT * FROM erp_navigation WHERE role IN ({roleList}) AND superior='0' AND state='0' ORDER BY sort ASC"
            Dim dt = DatabaseModule.ExecuteQuery(sql, MySQL_Auth)

            For Each row As DataRow In dt.Rows
                Dim nodeId = SafeString(row("id"))
                Dim nodeName = GBKToUTF8(SafeString(row("navigation")))
                Dim parentNode As New TreeNode(nodeName)
                parentNode.Tag = nodeId

                Dim childSql = $"SELECT * FROM erp_navigation WHERE role IN ({roleList}) AND superior='{nodeId}' AND state='0' ORDER BY sort ASC"
                Dim childDt = DatabaseModule.ExecuteQuery(childSql, MySQL_Auth)
                For Each childRow As DataRow In childDt.Rows
                    Dim childNode As New TreeNode(GBKToUTF8(SafeString(childRow("navigation"))))
                    childNode.Tag = SafeString(childRow("id"))
                    parentNode.Nodes.Add(childNode)
                Next
                treeView.Nodes.Add(parentNode)
            Next

            treeView.CollapseAll()
            treeView.SelectedNode = homeNode
        Catch ex As Exception
            ShowError("加载导航树失败：" & ex.Message)
        Finally
            isUiInitializing = False
        End Try
    End Sub

    Private Sub TreeView_AfterSelect(sender As Object, e As TreeViewEventArgs)
        If e.Node Is Nothing Then Return
        If isUiInitializing Then Return

        currentPageName = e.Node.Text
        currentPageID = SafeString(e.Node.Tag)

        If e.Node.Parent IsNot Nothing Then
            Dim parentNode = e.Node.Parent
            For Each node As TreeNode In treeView.Nodes
                If node IsNot parentNode AndAlso node.Nodes.Count > 0 Then
                    node.Collapse()
                End If
            Next
        End If

        RouteToPage(currentPageName, currentPageID)
    End Sub

    Private Sub RouteToPage(pageName As String, pageId As String)
        Select Case pageName
            Case "首页"
                ShowHomePage()
            Case "系统设置", "商铺信息"
                HideHomePanel()
                ShowSystemSettings()
            Case "账户列表", "用户管理"
                HideHomePanel()
                ShowAccountList()
            Case "岗位分组"
                HideHomePanel()
                ShowRoleGroupPage()
            Case "商品列表"
                HideHomePanel()
                ShowProductList()
            Case "商品入库"
                HideHomePanel()
                ShowInboundList()
            Case "商品销售"
                HideHomePanel()
                ShowSalesList()
            Case "商品调拨"
                HideHomePanel()
                ShowTransferList()
            Case "商品退库"
                HideHomePanel()
                ShowReturnList()
            Case "商品回收"
                HideHomePanel()
                ShowRecoveryList()
            Case "商品预售"
                HideHomePanel()
                ShowPresaleList()
            Case "商品退货"
                HideHomePanel()
                ShowRefundList()
            Case "会员列表"
                HideHomePanel()
                ShowMemberList()
            Case "实时库存"
                HideHomePanel()
                ShowInventoryList()
            Case "日志记录", "系统日志"
                Dim logForm As New SystemLogForm()
                logForm.ShowDialog()
            Case "登录日志"
                Dim loginLogForm As New LoginLogForm()
                loginLogForm.ShowDialog()
            Case "账户管理"
                Dim accountForm As New AccountForm()
                accountForm.ShowDialog()
            Case "个人信息"
                Dim personalForm As New PersonalInfoForm()
                personalForm.ShowDialog()
            Case "密码修改"
                Dim passwordForm As New PasswordChangeForm()
                passwordForm.ShowDialog()
            Case "品类管理"
                Dim categoryForm As New CategoryForm()
                categoryForm.ShowDialog()
            Case "规格管理"
                Dim specForm As New SpecForm()
                specForm.ShowDialog()
            Case "来源管理"
                Dim sourceForm As New SourceForm()
                sourceForm.ShowDialog()
            Case "结算方式"
                Dim settlementForm As New SettlementForm()
                settlementForm.ShowDialog()
            Case "品类属性"
                Dim categoryAttrForm As New CategoryAttrForm()
                categoryAttrForm.ShowDialog()
            Case "回收名称"
                Dim recoveryTitleForm As New RecoveryTitleForm()
                recoveryTitleForm.ShowDialog()
            Case "收支管理"
                Dim financeForm As New FinanceForm()
                financeForm.ShowDialog()
            Case "收支名称"
                Dim financeTitleForm As New FinanceTitleForm()
                financeTitleForm.ShowDialog()
            Case "收支卡号"
                Dim financeAccountForm As New FinanceAccountForm()
                financeAccountForm.ShowDialog()
            Case "结账结料"
                Dim factoryForm As New FactorySettlementForm()
                factoryForm.ShowDialog()
            Case "物资盘点"
                Dim inventoryCheckForm As New InventoryCheckForm()
                inventoryCheckForm.ShowDialog()
            Case Else
                HideHomePanel()
                dgvMain.DataSource = Nothing
                dgvMain.Rows.Clear()
                dgvMain.Columns.Clear()
                dgvMain.Columns.Add("info", "提示")
                dgvMain.Rows.Add($"页面 [{pageName}] (ID:{pageId}) 暂未实现")
        End Select
    End Sub

    Private Sub ShowHomePage()
        lblTitle.Text = "首页"
        dgvMain.Visible = False
        dgvMain.SendToBack()
        homePanel.Visible = True
        homePanel.BringToFront()
        panelPage.Visible = False
        panelContent.BackColor = BgColor
        lblHomeWelcome.Text = $"欢迎回来，{UserName}{If(String.IsNullOrEmpty(UserPostName), "", UserPostName)}"
        LayoutHomeDashboard()
        panelContent.PerformLayout()
        homePanel.PerformLayout()
        dashboardTimer.Start()
        RefreshDashboardData()
    End Sub

    Private Sub HideHomePanel()
        If dashboardTimer IsNot Nothing Then dashboardTimer.Stop()
        If homePanel IsNot Nothing Then
            homePanel.Visible = False
            homePanel.SendToBack()
        End If
        dgvMain.Visible = True
        dgvMain.BringToFront()
        panelPage.Visible = True
        panelContent.BackColor = CardColor
    End Sub

    Private Sub DashboardTimer_Tick(sender As Object, e As EventArgs)
        If homePanel.Visible Then RefreshDashboardData()
    End Sub

    Private Sub RefreshDashboardData()
        If isDashboardLoading Then Return
        isDashboardLoading = True

        Task.Run(Sub()
                     Try
                         Dim snapshot = HomeDashboardService.LoadDashboardSnapshot(trendPoints)
                         BeginInvoke(New Action(Sub()
                                                    ApplyDashboardSnapshot(snapshot)
                                                    isDashboardLoading = False
                                                End Sub))
                     Catch ex As Exception
                         BeginInvoke(New Action(Sub()
                                                    isDashboardLoading = False
                                                End Sub))
                     End Try
                 End Sub)
    End Sub

    Private Sub ApplyDashboardSnapshot(snapshot As HomeDashboardService.DashboardSnapshot)
        trendPoints = snapshot.Trend
        Dim kpi = snapshot.Kpi
        Dim mask = hideValuesEnabled

        Dim values = {
            FormatDashboardValue(kpi.SalesAmount, "0.00", mask),
            FormatDashboardValue(kpi.SalesQty, "0.00", mask),
            FormatDashboardValue(kpi.SalesWeight, "0.000", mask),
            FormatDashboardValue(kpi.RecoveryAmount, "0.00", mask),
            FormatDashboardValue(kpi.RecoveryWeight, "0.000", mask),
            FormatDashboardValue(kpi.ReceivedAmount, "0.00", mask),
            FormatDashboardValue(kpi.ActualWeight, "0.000", mask)
        }

        For i As Integer = 0 To lblKpiValues.Length - 1
            lblKpiValues(i).Text = values(i)
        Next

        UpdateTrendChart(snapshot)
        UpdateRankingGrid(snapshot.Ranking, mask)
        UpdatePieChart(snapshot.Categories)
    End Sub

    Private Function FormatDashboardValue(value As Decimal, format As String, mask As Boolean) As String
        If mask Then Return "***"
        Return value.ToString(format)
    End Function

    Private Sub UpdateTrendChart(snapshot As HomeDashboardService.DashboardSnapshot)
        chartTrend.Series("实收").Points.Clear()
        chartTrend.Series("销售").Points.Clear()
        chartTrend.Series("回收").Points.Clear()

        Dim area = chartTrend.ChartAreas(0)
        area.AxisY.Maximum = CDbl(snapshot.TrendMax)

        For Each point In snapshot.Trend
            Dim x = point.Label
            Dim received = If(hideValuesEnabled, 0D, point.Received)
            Dim sales = If(hideValuesEnabled, 0D, point.Sales)
            Dim recovery = If(hideValuesEnabled, 0D, point.Recovery)
            chartTrend.Series("实收").Points.AddXY(x, received)
            chartTrend.Series("销售").Points.AddXY(x, sales)
            chartTrend.Series("回收").Points.AddXY(x, recovery)
        Next
    End Sub

    Private Sub UpdateRankingGrid(rows As List(Of HomeDashboardService.SalesRankRow), mask As Boolean)
        dgvRanking.Rows.Clear()
        For Each row In rows
            dgvRanking.Rows.Add(
                row.Rank,
                row.ItemNumber,
                row.Category,
                row.Spec,
                If(mask, "***", row.Quantity),
                If(mask, "***", row.Amount)
            )
        Next
    End Sub

    Private Sub UpdatePieChart(slices As List(Of HomeDashboardService.CategorySlice))
        chartPie.Series("Category").Points.Clear()
        panelPieLegend.Controls.Clear()

        If slices.Count = 0 Then
            chartPie.Series("Category").Points.AddXY("暂无数据", 1)
            chartPie.Series("Category").Points(0).Color = Color.FromArgb(230, 230, 230)
            Return
        End If

        For Each slice In slices
            Dim idx = chartPie.Series("Category").Points.AddXY(slice.Name, slice.Weight)
            chartPie.Series("Category").Points(idx).Color = slice.Color

            Dim legendItem As New Panel()
            legendItem.Size = New Size(140, 24)
            legendItem.BackColor = Color.Transparent
            Dim colorBox As New Panel()
            colorBox.Size = New Size(10, 10)
            colorBox.Location = New Point(0, 6)
            colorBox.BackColor = slice.Color
            legendItem.Controls.Add(colorBox)
            Dim text As New Label()
            text.Text = $"{slice.Name} {slice.Weight:0.000}g"
            text.Location = New Point(16, 2)
            text.AutoSize = True
            text.ForeColor = TextDark
            text.Font = New Font("微软雅黑", 8.5F)
            legendItem.Controls.Add(text)
            panelPieLegend.Controls.Add(legendItem)
        Next
    End Sub

    Private Sub ShowSystemSettings()
        lblTitle.Text = "系统设置"
        ' 打开系统设置窗口
        Dim settingsForm As New SystemSettingsForm()
        settingsForm.ShowDialog()
    End Sub

    Private Sub ShowAccountList()
        lblTitle.Text = "账户列表"
        currentTableName = "xipunum_erp_user"
        LoadData()
    End Sub

    Private Sub ShowRoleGroupPage()
        lblTitle.Text = "岗位分组"
        currentTableName = "xipunum_erp_type"
        LoadData()
    End Sub

    Private Sub ShowProductList()
        lblTitle.Text = "商品列表"
        currentTableName = "xipunum_erp_shop"
        LoadData()
    End Sub

    Private Sub ShowInboundList()
        lblTitle.Text = "商品入库"
        currentTableName = "xipunum_erp_store_order"
        LoadData()
    End Sub

    Private Sub ShowSalesList()
        lblTitle.Text = "商品销售"
        currentTableName = "xipunum_erp_outbound_order"
        LoadData()
    End Sub

    Private Sub ShowTransferList()
        lblTitle.Text = "商品调拨"
        currentTableName = "xipunum_erp_transfer_order"
        LoadData()
    End Sub

    Private Sub ShowReturnList()
        lblTitle.Text = "商品退库"
        currentTableName = "xipunum_erp_tuiku_order"
        LoadData()
    End Sub

    Private Sub ShowRecoveryList()
        lblTitle.Text = "商品回收"
        currentTableName = "xipunum_erp_retreat_order"
        LoadData()
    End Sub

    Private Sub ShowPresaleList()
        lblTitle.Text = "商品预售"
        currentTableName = "xipunum_erp_presale_order"
        LoadData()
    End Sub

    Private Sub ShowRefundList()
        lblTitle.Text = "商品退货"
        currentTableName = "xipunum_erp_return_order"
        LoadData()
    End Sub

    Private Sub ShowMemberList()
        lblTitle.Text = "会员列表"
        currentTableName = "xipunum_erp_member"
        LoadData()
    End Sub

    Private Sub ShowInventoryList()
        lblTitle.Text = "实时库存"
        ' 使用关联查询获取库存信息
        Try
            Dim shopPermission As String = UserShopPermission
            If String.IsNullOrEmpty(shopPermission) Then shopPermission = "-1"

            Dim sql = $"SELECT a.poduct_code, a.product_name, a.item_number AS kuanhao, " &
                     $"b.title AS pinlei, c.title AS guige, a.caizhi, a.single AS danzhong, " &
                     $"CAST(COALESCE(SUM(d.quantity),0) AS DECIMAL(30,2)) AS kucunshuliang, " &
                     $"CAST(COALESCE(SUM(d.jinzhong),0) AS DECIMAL(30,3)) AS kucunjinzhong, " &
                     $"e.title AS kufang " &
                     $"FROM xipunum_erp_shop AS a " &
                     $"INNER JOIN xipunum_erp_shop_kucun AS d ON d.poduct_code = a.poduct_code " &
                     $"LEFT JOIN xipunum_erp_category AS b ON b.id = a.category_id " &
                     $"LEFT JOIN xipunum_erp_specs AS c ON c.id = a.specification_id " &
                     $"LEFT JOIN xipunum_erp_type AS e ON e.id = d.kufang " &
                     $"WHERE (d.quantity > 0 OR d.jinzhong > 0) AND d.kufang IN ({shopPermission}) " &
                     $"GROUP BY a.poduct_code, d.kufang ORDER BY a.id DESC LIMIT 1000"

            Dim dt = DatabaseModule.ExecuteQuery(sql)
            dgvMain.DataSource = dt
            lblPageInfo.Text = $"共 {dt.Rows.Count} 条记录"
            HideHomePanel()
        Catch ex As Exception
            ShowError("加载库存数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub ShowLogPage()
        lblTitle.Text = "日志记录"
        currentTableName = "xipunum_erp_xitong_log"
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            If String.IsNullOrEmpty(currentTableName) Then Return
            dgvMain.Visible = True
            HideHomePanel()

            ' 计算总记录数
            Dim countSql = $"SELECT COUNT(*) FROM {currentTableName}"
            Dim countResult = DatabaseModule.ExecuteScalar(countSql)
            totalRecords = If(countResult IsNot Nothing, Convert.ToInt32(countResult), 0)
            totalPages = Math.Max(1, CInt(Math.Ceiling(totalRecords / CDbl(pageSize))))

            ' 确保当前页不超出范围
            If currentPage > totalPages Then currentPage = totalPages
            If currentPage < 1 Then currentPage = 1

            ' 查询当前页数据
            Dim sql = $"SELECT * FROM {currentTableName} ORDER BY {currentSortField} {currentSortOrder} LIMIT {(currentPage - 1) * pageSize},{pageSize}"
            Dim dt = DatabaseModule.ExecuteQuery(sql)
            dgvMain.DataSource = dt

            ' 更新分页信息
            lblPageInfo.Text = $"共 {totalRecords} 条记录，第 {currentPage}/{totalPages} 页"
            btnFirst.Enabled = (currentPage > 1)
            btnPrev.Enabled = (currentPage > 1)
            btnNext.Enabled = (currentPage < totalPages)
            btnLast.Enabled = (currentPage < totalPages)
        Catch ex As Exception
            ShowError("加载数据失败：" & ex.Message)
        End Try
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        currentPage = 1
        LoadData()
    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        If currentPage > 1 Then
            currentPage -= 1
            LoadData()
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If currentPage < totalPages Then
            currentPage += 1
            LoadData()
        End If
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        currentPage = totalPages
        LoadData()
    End Sub

End Class
