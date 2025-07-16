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
        Me.cmdExtractToDirectory = New System.Windows.Forms.Button()
        Me.lstVerRec = New System.Windows.Forms.ListBox()
        Me.lstNa = New System.Windows.Forms.ListBox()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdExtractToDirectory
        '
        Me.cmdExtractToDirectory.Location = New System.Drawing.Point(42, 33)
        Me.cmdExtractToDirectory.Name = "cmdExtractToDirectory"
        Me.cmdExtractToDirectory.Size = New System.Drawing.Size(135, 37)
        Me.cmdExtractToDirectory.TabIndex = 0
        Me.cmdExtractToDirectory.Text = "ExtractToDirectory"
        Me.cmdExtractToDirectory.UseVisualStyleBackColor = True
        '
        'lstVerRec
        '
        Me.lstVerRec.FormattingEnabled = True
        Me.lstVerRec.ItemHeight = 12
        Me.lstVerRec.Location = New System.Drawing.Point(42, 87)
        Me.lstVerRec.Name = "lstVerRec"
        Me.lstVerRec.Size = New System.Drawing.Size(351, 328)
        Me.lstVerRec.TabIndex = 1
        '
        'lstNa
        '
        Me.lstNa.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstNa.Font = New System.Drawing.Font("ＭＳ ゴシック", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lstNa.FormattingEnabled = True
        Me.lstNa.HorizontalScrollbar = True
        Me.lstNa.ItemHeight = 19
        Me.lstNa.Location = New System.Drawing.Point(412, 87)
        Me.lstNa.Name = "lstNa"
        Me.lstNa.Size = New System.Drawing.Size(728, 365)
        Me.lstNa.TabIndex = 2
        '
        'lblCount
        '
        Me.lblCount.AutoSize = True
        Me.lblCount.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCount.Location = New System.Drawing.Point(410, 58)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(0, 16)
        Me.lblCount.TabIndex = 3
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1168, 502)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.lstNa)
        Me.Controls.Add(Me.lstVerRec)
        Me.Controls.Add(Me.cmdExtractToDirectory)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdExtractToDirectory As Button
    Friend WithEvents lstVerRec As ListBox
    Friend WithEvents lstNa As ListBox
    Friend WithEvents lblCount As Label
End Class
