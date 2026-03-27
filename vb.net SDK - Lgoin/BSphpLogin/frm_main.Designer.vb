<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_main
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.button7 = New System.Windows.Forms.Button()
        Me.button6 = New System.Windows.Forms.Button()
        Me.button4 = New System.Windows.Forms.Button()
        Me.button5 = New System.Windows.Forms.Button()
        Me.button2 = New System.Windows.Forms.Button()
        Me.button3 = New System.Windows.Forms.Button()
        Me.button1 = New System.Windows.Forms.Button()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.label1 = New System.Windows.Forms.Label()
        Me.groupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'button7
        '
        Me.button7.Location = New System.Drawing.Point(30, 277)
        Me.button7.Name = "button7"
        Me.button7.Size = New System.Drawing.Size(116, 32)
        Me.button7.TabIndex = 10
        Me.button7.Text = "取软件配置信息段"
        Me.button7.UseVisualStyleBackColor = True
        '
        'button6
        '
        Me.button6.Location = New System.Drawing.Point(338, 229)
        Me.button6.Name = "button6"
        Me.button6.Size = New System.Drawing.Size(116, 33)
        Me.button6.TabIndex = 4
        Me.button6.Text = "登陆状态"
        Me.button6.UseVisualStyleBackColor = True
        '
        'button4
        '
        Me.button4.Location = New System.Drawing.Point(189, 229)
        Me.button4.Name = "button4"
        Me.button4.Size = New System.Drawing.Size(116, 33)
        Me.button4.TabIndex = 5
        Me.button4.Text = "取验证数据"
        Me.button4.UseVisualStyleBackColor = True
        '
        'button5
        '
        Me.button5.Location = New System.Drawing.Point(338, 174)
        Me.button5.Name = "button5"
        Me.button5.Size = New System.Drawing.Size(116, 33)
        Me.button5.TabIndex = 6
        Me.button5.Text = "取绑定特征"
        Me.button5.UseVisualStyleBackColor = True
        '
        'button2
        '
        Me.button2.Location = New System.Drawing.Point(189, 174)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(116, 33)
        Me.button2.TabIndex = 7
        Me.button2.Text = "到期时间"
        Me.button2.UseVisualStyleBackColor = True
        '
        'button3
        '
        Me.button3.Location = New System.Drawing.Point(30, 229)
        Me.button3.Name = "button3"
        Me.button3.Size = New System.Drawing.Size(116, 33)
        Me.button3.TabIndex = 8
        Me.button3.Text = "取用户信息"
        Me.button3.UseVisualStyleBackColor = True
        '
        'button1
        '
        Me.button1.Location = New System.Drawing.Point(30, 174)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(116, 33)
        Me.button1.TabIndex = 9
        Me.button1.Text = "连接测试"
        Me.button1.UseVisualStyleBackColor = True
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Location = New System.Drawing.Point(12, 12)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(455, 124)
        Me.groupBox1.TabIndex = 3
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "软件描述"
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(20, 24)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(424, 90)
        Me.label1.TabIndex = 0
        Me.label1.Text = "介绍内容"
        '
        'frm_main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 318)
        Me.Controls.Add(Me.button7)
        Me.Controls.Add(Me.button6)
        Me.Controls.Add(Me.button4)
        Me.Controls.Add(Me.button5)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.button3)
        Me.Controls.Add(Me.button1)
        Me.Controls.Add(Me.groupBox1)
        Me.Name = "frm_main"
        Me.Text = "程序主界面"
        Me.groupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents button7 As Button
    Private WithEvents button6 As Button
    Private WithEvents button4 As Button
    Private WithEvents button5 As Button
    Private WithEvents button2 As Button
    Private WithEvents button3 As Button
    Private WithEvents button1 As Button
    Private WithEvents groupBox1 As GroupBox
    Private WithEvents label1 As Label
End Class
