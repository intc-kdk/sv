<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ファイルToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSMI_MM_手順書をリセット = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSMI_MM_DBのバックアップ = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSMI_MM_終了 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSMI_MM_編集 = New System.Windows.Forms.ToolStripMenuItem()
        Me.管理補助ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSMI_MM_ログに追記 = New System.Windows.Forms.ToolStripMenuItem()
        Me.テストメニューToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.発呼実験ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.雑ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.TSMI_TTM_ウインドウを開く = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSMI_TTM_手順書をリセット = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSMI_TTM_DBのバックアップ = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSMI_TTM_編集 = New System.Windows.Forms.ToolStripMenuItem()
        Me.管理補助ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSMI_TTM_ログに追記 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TSS_TTM_編集 = New System.Windows.Forms.ToolStripSeparator()
        Me.TSMI_TTM_終了 = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TB_設定 = New System.Windows.Forms.TextBox()
        Me.MenuStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ファイルToolStripMenuItem, Me.TSMI_MM_編集, Me.テストメニューToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(387, 26)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ファイルToolStripMenuItem
        '
        Me.ファイルToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSMI_MM_手順書をリセット, Me.TSMI_MM_DBのバックアップ, Me.ToolStripSeparator4, Me.TSMI_MM_終了})
        Me.ファイルToolStripMenuItem.Name = "ファイルToolStripMenuItem"
        Me.ファイルToolStripMenuItem.Size = New System.Drawing.Size(68, 22)
        Me.ファイルToolStripMenuItem.Text = "ファイル"
        '
        'TSMI_MM_手順書をリセット
        '
        Me.TSMI_MM_手順書をリセット.Name = "TSMI_MM_手順書をリセット"
        Me.TSMI_MM_手順書をリセット.Size = New System.Drawing.Size(184, 22)
        Me.TSMI_MM_手順書をリセット.Text = "手順書を初期化する"
        '
        'TSMI_MM_DBのバックアップ
        '
        Me.TSMI_MM_DBのバックアップ.Name = "TSMI_MM_DBのバックアップ"
        Me.TSMI_MM_DBのバックアップ.Size = New System.Drawing.Size(184, 22)
        Me.TSMI_MM_DBのバックアップ.Text = "DBのバックアップ"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(181, 6)
        '
        'TSMI_MM_終了
        '
        Me.TSMI_MM_終了.Name = "TSMI_MM_終了"
        Me.TSMI_MM_終了.Size = New System.Drawing.Size(184, 22)
        Me.TSMI_MM_終了.Text = "終了"
        '
        'TSMI_MM_編集
        '
        Me.TSMI_MM_編集.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.管理補助ToolStripMenuItem})
        Me.TSMI_MM_編集.Name = "TSMI_MM_編集"
        Me.TSMI_MM_編集.Size = New System.Drawing.Size(44, 22)
        Me.TSMI_MM_編集.Text = "編集"
        '
        '管理補助ToolStripMenuItem
        '
        Me.管理補助ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSMI_MM_ログに追記})
        Me.管理補助ToolStripMenuItem.Name = "管理補助ToolStripMenuItem"
        Me.管理補助ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.管理補助ToolStripMenuItem.Text = "管理補助"
        '
        'TSMI_MM_ログに追記
        '
        Me.TSMI_MM_ログに追記.Name = "TSMI_MM_ログに追記"
        Me.TSMI_MM_ログに追記.Size = New System.Drawing.Size(136, 22)
        Me.TSMI_MM_ログに追記.Text = "ログに追記"
        '
        'テストメニューToolStripMenuItem
        '
        Me.テストメニューToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.発呼実験ToolStripMenuItem1, Me.雑ToolStripMenuItem})
        Me.テストメニューToolStripMenuItem.Name = "テストメニューToolStripMenuItem"
        Me.テストメニューToolStripMenuItem.Size = New System.Drawing.Size(104, 22)
        Me.テストメニューToolStripMenuItem.Text = "テストメニュー"
        Me.テストメニューToolStripMenuItem.Visible = False
        '
        '発呼実験ToolStripMenuItem1
        '
        Me.発呼実験ToolStripMenuItem1.Name = "発呼実験ToolStripMenuItem1"
        Me.発呼実験ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
        Me.発呼実験ToolStripMenuItem1.Text = "発呼実験"
        Me.発呼実験ToolStripMenuItem1.ToolTipText = "「99@(現在日時)$」を送信します"
        '
        '雑ToolStripMenuItem
        '
        Me.雑ToolStripMenuItem.Name = "雑ToolStripMenuItem"
        Me.雑ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.雑ToolStripMenuItem.Text = "雑"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSMI_TTM_ウインドウを開く, Me.ToolStripSeparator3, Me.TSMI_TTM_手順書をリセット, Me.TSMI_TTM_DBのバックアップ, Me.ToolStripSeparator1, Me.TSMI_TTM_編集, Me.TSS_TTM_編集, Me.TSMI_TTM_終了})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(178, 132)
        '
        'TSMI_TTM_ウインドウを開く
        '
        Me.TSMI_TTM_ウインドウを開く.Name = "TSMI_TTM_ウインドウを開く"
        Me.TSMI_TTM_ウインドウを開く.Size = New System.Drawing.Size(177, 22)
        Me.TSMI_TTM_ウインドウを開く.Text = "ウインドウを開く"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(174, 6)
        '
        'TSMI_TTM_手順書をリセット
        '
        Me.TSMI_TTM_手順書をリセット.Name = "TSMI_TTM_手順書をリセット"
        Me.TSMI_TTM_手順書をリセット.Size = New System.Drawing.Size(177, 22)
        Me.TSMI_TTM_手順書をリセット.Text = "手順書をリセット"
        '
        'TSMI_TTM_DBのバックアップ
        '
        Me.TSMI_TTM_DBのバックアップ.Name = "TSMI_TTM_DBのバックアップ"
        Me.TSMI_TTM_DBのバックアップ.Size = New System.Drawing.Size(177, 22)
        Me.TSMI_TTM_DBのバックアップ.Text = "DBのバックアップ"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(174, 6)
        '
        'TSMI_TTM_編集
        '
        Me.TSMI_TTM_編集.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.管理補助ToolStripMenuItem1})
        Me.TSMI_TTM_編集.Name = "TSMI_TTM_編集"
        Me.TSMI_TTM_編集.Size = New System.Drawing.Size(177, 22)
        Me.TSMI_TTM_編集.Text = "編集"
        '
        '管理補助ToolStripMenuItem1
        '
        Me.管理補助ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSMI_TTM_ログに追記})
        Me.管理補助ToolStripMenuItem1.Name = "管理補助ToolStripMenuItem1"
        Me.管理補助ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
        Me.管理補助ToolStripMenuItem1.Text = "管理補助"
        '
        'TSMI_TTM_ログに追記
        '
        Me.TSMI_TTM_ログに追記.Name = "TSMI_TTM_ログに追記"
        Me.TSMI_TTM_ログに追記.Size = New System.Drawing.Size(136, 22)
        Me.TSMI_TTM_ログに追記.Text = "ログに追記"
        '
        'TSS_TTM_編集
        '
        Me.TSS_TTM_編集.Name = "TSS_TTM_編集"
        Me.TSS_TTM_編集.Size = New System.Drawing.Size(174, 6)
        '
        'TSMI_TTM_終了
        '
        Me.TSMI_TTM_終了.Name = "TSMI_TTM_終了"
        Me.TSMI_TTM_終了.Size = New System.Drawing.Size(177, 22)
        Me.TSMI_TTM_終了.Text = "終了"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "タブレットサーバ"
        Me.NotifyIcon1.Visible = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 26)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(387, 310)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TB_設定)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(379, 284)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "設定"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TB_設定
        '
        Me.TB_設定.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TB_設定.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TB_設定.Location = New System.Drawing.Point(3, 3)
        Me.TB_設定.Multiline = True
        Me.TB_設定.Name = "TB_設定"
        Me.TB_設定.ReadOnly = True
        Me.TB_設定.Size = New System.Drawing.Size(373, 278)
        Me.TB_設定.TabIndex = 0
        Me.TB_設定.WordWrap = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(387, 336)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.ShowInTaskbar = False
        Me.Text = "タブレットサーバ"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ファイルToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_MM_終了 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents テストメニューToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 発呼実験ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents TSMI_TTM_ウインドウを開く As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TSMI_TTM_終了 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TB_設定 As System.Windows.Forms.TextBox
    Friend WithEvents 雑ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_MM_手順書をリセット As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TSMI_TTM_手順書をリセット As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_TTM_DBのバックアップ As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_MM_DBのバックアップ As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TSMI_MM_編集 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 管理補助ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_MM_ログに追記 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_TTM_編集 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 管理補助ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSMI_TTM_ログに追記 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TSS_TTM_編集 As System.Windows.Forms.ToolStripSeparator

End Class
