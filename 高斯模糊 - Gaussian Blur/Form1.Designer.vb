<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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

    '注意:  以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Test1 = New System.Windows.Forms.Button()
        Me.Test2 = New System.Windows.Forms.Button()
        Me.Test3 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SetPixel = New System.Windows.Forms.NumericUpDown()
        Me.Test4 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.SetPixel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Test1
        '
        Me.Test1.ForeColor = System.Drawing.Color.Red
        Me.Test1.Image = CType(resources.GetObject("Test1.Image"), System.Drawing.Image)
        Me.Test1.Location = New System.Drawing.Point(547, 12)
        Me.Test1.Name = "Test1"
        Me.Test1.Size = New System.Drawing.Size(75, 33)
        Me.Test1.TabIndex = 1
        Me.Test1.Text = "Test-1"
        Me.Test1.UseVisualStyleBackColor = True
        '
        'Test2
        '
        Me.Test2.ForeColor = System.Drawing.Color.Red
        Me.Test2.Image = CType(resources.GetObject("Test2.Image"), System.Drawing.Image)
        Me.Test2.Location = New System.Drawing.Point(547, 51)
        Me.Test2.Name = "Test2"
        Me.Test2.Size = New System.Drawing.Size(75, 33)
        Me.Test2.TabIndex = 2
        Me.Test2.Text = "Test-2"
        Me.Test2.UseVisualStyleBackColor = True
        '
        'Test3
        '
        Me.Test3.ForeColor = System.Drawing.Color.Red
        Me.Test3.Image = CType(resources.GetObject("Test3.Image"), System.Drawing.Image)
        Me.Test3.Location = New System.Drawing.Point(547, 90)
        Me.Test3.Name = "Test3"
        Me.Test3.Size = New System.Drawing.Size(75, 33)
        Me.Test3.TabIndex = 3
        Me.Test3.Text = "Test-3"
        Me.Test3.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(547, 185)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "模糊范围："
        '
        'SetPixel
        '
        Me.SetPixel.Location = New System.Drawing.Point(547, 200)
        Me.SetPixel.Name = "SetPixel"
        Me.SetPixel.Size = New System.Drawing.Size(58, 21)
        Me.SetPixel.TabIndex = 6
        Me.SetPixel.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Test4
        '
        Me.Test4.ForeColor = System.Drawing.Color.Red
        Me.Test4.Image = CType(resources.GetObject("Test4.Image"), System.Drawing.Image)
        Me.Test4.Location = New System.Drawing.Point(547, 129)
        Me.Test4.Name = "Test4"
        Me.Test4.Size = New System.Drawing.Size(75, 33)
        Me.Test4.TabIndex = 7
        Me.Test4.Text = "Test-4"
        Me.Test4.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(547, 227)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(1, 1)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(540, 360)
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(626, 363)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Test4)
        Me.Controls.Add(Me.SetPixel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Test3)
        Me.Controls.Add(Me.Test2)
        Me.Controls.Add(Me.Test1)
        Me.Name = "Form1"
        Me.Text = "高斯模糊 - 第一版 无优化"
        CType(Me.SetPixel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Test1 As System.Windows.Forms.Button
    Friend WithEvents Test2 As System.Windows.Forms.Button
    Friend WithEvents Test3 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SetPixel As System.Windows.Forms.NumericUpDown
    Friend WithEvents Test4 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox

End Class
