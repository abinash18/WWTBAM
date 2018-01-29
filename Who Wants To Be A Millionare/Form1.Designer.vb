<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.mnuMenuStripOne = New System.Windows.Forms.MenuStrip()
        Me.mnuProgram = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNewGame = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWalkAway = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLifeLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpAndInstructions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnAnswerChoiceA = New System.Windows.Forms.Button()
        Me.btnAnswerChoiceD = New System.Windows.Forms.Button()
        Me.btnAnswerChoiceB = New System.Windows.Forms.Button()
        Me.btnAnswerChoiceC = New System.Windows.Forms.Button()
        Me.lblOptionA = New System.Windows.Forms.Label()
        Me.lblOptionC = New System.Windows.Forms.Label()
        Me.lblOptionD = New System.Windows.Forms.Label()
        Me.lblOptionB = New System.Windows.Forms.Label()
        Me.picHintPicture = New System.Windows.Forms.PictureBox()
        Me.lstMoneyWon = New System.Windows.Forms.ListBox()
        Me.btnLifeLine = New System.Windows.Forms.Button()
        Me.btnWalkAway = New System.Windows.Forms.Button()
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.lblTimer = New System.Windows.Forms.Label()
        Me.tmrQuestionTimer = New System.Windows.Forms.Timer(Me.components)
        Me.btnStartNewGame = New System.Windows.Forms.Button()
        Me.prgTimeLeft = New System.Windows.Forms.ProgressBar()
        Me.prgTimeRight = New System.Windows.Forms.ProgressBar()
        Me.tmrEndGameCelebration = New System.Windows.Forms.Timer(Me.components)
        Me.mnuMenuStripOne.SuspendLayout()
        CType(Me.picHintPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mnuMenuStripOne
        '
        Me.mnuMenuStripOne.BackColor = System.Drawing.SystemColors.Control
        Me.mnuMenuStripOne.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuProgram})
        Me.mnuMenuStripOne.Location = New System.Drawing.Point(0, 0)
        Me.mnuMenuStripOne.Name = "mnuMenuStripOne"
        Me.mnuMenuStripOne.Size = New System.Drawing.Size(695, 24)
        Me.mnuMenuStripOne.TabIndex = 0
        Me.mnuMenuStripOne.Text = "MenuStrip1"
        '
        'mnuProgram
        '
        Me.mnuProgram.BackColor = System.Drawing.SystemColors.Control
        Me.mnuProgram.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNewGame, Me.mnuWalkAway, Me.mnuLifeLine, Me.mnuHelpAndInstructions, Me.mnuExit})
        Me.mnuProgram.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.mnuProgram.Name = "mnuProgram"
        Me.mnuProgram.Size = New System.Drawing.Size(65, 20)
        Me.mnuProgram.Text = "Program"
        '
        'mnuNewGame
        '
        Me.mnuNewGame.Name = "mnuNewGame"
        Me.mnuNewGame.Size = New System.Drawing.Size(189, 22)
        Me.mnuNewGame.Text = "New Game"
        '
        'mnuWalkAway
        '
        Me.mnuWalkAway.Name = "mnuWalkAway"
        Me.mnuWalkAway.Size = New System.Drawing.Size(189, 22)
        Me.mnuWalkAway.Text = "Walk Away"
        Me.mnuWalkAway.Visible = False
        '
        'mnuLifeLine
        '
        Me.mnuLifeLine.Name = "mnuLifeLine"
        Me.mnuLifeLine.Size = New System.Drawing.Size(189, 22)
        Me.mnuLifeLine.Text = "Life Line"
        Me.mnuLifeLine.Visible = False
        '
        'mnuHelpAndInstructions
        '
        Me.mnuHelpAndInstructions.Name = "mnuHelpAndInstructions"
        Me.mnuHelpAndInstructions.Size = New System.Drawing.Size(189, 22)
        Me.mnuHelpAndInstructions.Text = "Help And Instructions"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(189, 22)
        Me.mnuExit.Text = "Exit"
        '
        'btnAnswerChoiceA
        '
        Me.btnAnswerChoiceA.BackColor = System.Drawing.Color.White
        Me.btnAnswerChoiceA.ForeColor = System.Drawing.Color.Black
        Me.btnAnswerChoiceA.Location = New System.Drawing.Point(64, 400)
        Me.btnAnswerChoiceA.Name = "btnAnswerChoiceA"
        Me.btnAnswerChoiceA.Size = New System.Drawing.Size(272, 48)
        Me.btnAnswerChoiceA.TabIndex = 1
        Me.btnAnswerChoiceA.Tag = "a"
        Me.btnAnswerChoiceA.UseVisualStyleBackColor = False
        Me.btnAnswerChoiceA.Visible = False
        '
        'btnAnswerChoiceD
        '
        Me.btnAnswerChoiceD.BackColor = System.Drawing.Color.White
        Me.btnAnswerChoiceD.ForeColor = System.Drawing.Color.Black
        Me.btnAnswerChoiceD.Location = New System.Drawing.Point(384, 464)
        Me.btnAnswerChoiceD.Name = "btnAnswerChoiceD"
        Me.btnAnswerChoiceD.Size = New System.Drawing.Size(272, 48)
        Me.btnAnswerChoiceD.TabIndex = 2
        Me.btnAnswerChoiceD.Tag = "d"
        Me.btnAnswerChoiceD.UseVisualStyleBackColor = False
        Me.btnAnswerChoiceD.Visible = False
        '
        'btnAnswerChoiceB
        '
        Me.btnAnswerChoiceB.BackColor = System.Drawing.Color.White
        Me.btnAnswerChoiceB.ForeColor = System.Drawing.Color.Black
        Me.btnAnswerChoiceB.Location = New System.Drawing.Point(384, 400)
        Me.btnAnswerChoiceB.Name = "btnAnswerChoiceB"
        Me.btnAnswerChoiceB.Size = New System.Drawing.Size(272, 48)
        Me.btnAnswerChoiceB.TabIndex = 3
        Me.btnAnswerChoiceB.Tag = "b"
        Me.btnAnswerChoiceB.UseVisualStyleBackColor = False
        Me.btnAnswerChoiceB.Visible = False
        '
        'btnAnswerChoiceC
        '
        Me.btnAnswerChoiceC.BackColor = System.Drawing.Color.White
        Me.btnAnswerChoiceC.ForeColor = System.Drawing.Color.Black
        Me.btnAnswerChoiceC.Location = New System.Drawing.Point(64, 464)
        Me.btnAnswerChoiceC.Name = "btnAnswerChoiceC"
        Me.btnAnswerChoiceC.Size = New System.Drawing.Size(272, 48)
        Me.btnAnswerChoiceC.TabIndex = 4
        Me.btnAnswerChoiceC.Tag = "c"
        Me.btnAnswerChoiceC.UseVisualStyleBackColor = False
        Me.btnAnswerChoiceC.Visible = False
        '
        'lblOptionA
        '
        Me.lblOptionA.BackColor = System.Drawing.Color.Transparent
        Me.lblOptionA.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOptionA.ForeColor = System.Drawing.Color.Yellow
        Me.lblOptionA.Location = New System.Drawing.Point(16, 408)
        Me.lblOptionA.Name = "lblOptionA"
        Me.lblOptionA.Size = New System.Drawing.Size(40, 23)
        Me.lblOptionA.TabIndex = 5
        Me.lblOptionA.Tag = ""
        Me.lblOptionA.Text = "A:"
        Me.lblOptionA.Visible = False
        '
        'lblOptionC
        '
        Me.lblOptionC.BackColor = System.Drawing.Color.Transparent
        Me.lblOptionC.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOptionC.ForeColor = System.Drawing.Color.Yellow
        Me.lblOptionC.Location = New System.Drawing.Point(16, 472)
        Me.lblOptionC.Name = "lblOptionC"
        Me.lblOptionC.Size = New System.Drawing.Size(40, 23)
        Me.lblOptionC.TabIndex = 6
        Me.lblOptionC.Tag = ""
        Me.lblOptionC.Text = "C:"
        Me.lblOptionC.Visible = False
        '
        'lblOptionD
        '
        Me.lblOptionD.BackColor = System.Drawing.Color.Transparent
        Me.lblOptionD.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOptionD.ForeColor = System.Drawing.Color.Yellow
        Me.lblOptionD.Location = New System.Drawing.Point(344, 472)
        Me.lblOptionD.Name = "lblOptionD"
        Me.lblOptionD.Size = New System.Drawing.Size(40, 23)
        Me.lblOptionD.TabIndex = 7
        Me.lblOptionD.Tag = ""
        Me.lblOptionD.Text = "D:"
        Me.lblOptionD.Visible = False
        '
        'lblOptionB
        '
        Me.lblOptionB.BackColor = System.Drawing.Color.Transparent
        Me.lblOptionB.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOptionB.ForeColor = System.Drawing.Color.Yellow
        Me.lblOptionB.Location = New System.Drawing.Point(344, 408)
        Me.lblOptionB.Name = "lblOptionB"
        Me.lblOptionB.Size = New System.Drawing.Size(40, 23)
        Me.lblOptionB.TabIndex = 8
        Me.lblOptionB.Tag = ""
        Me.lblOptionB.Text = "B:"
        Me.lblOptionB.Visible = False
        '
        'picHintPicture
        '
        Me.picHintPicture.BackColor = System.Drawing.Color.Transparent
        Me.picHintPicture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.picHintPicture.Location = New System.Drawing.Point(184, 32)
        Me.picHintPicture.Name = "picHintPicture"
        Me.picHintPicture.Size = New System.Drawing.Size(336, 208)
        Me.picHintPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picHintPicture.TabIndex = 9
        Me.picHintPicture.TabStop = False
        Me.picHintPicture.Visible = False
        '
        'lstMoneyWon
        '
        Me.lstMoneyWon.BackColor = System.Drawing.SystemColors.Control
        Me.lstMoneyWon.Enabled = False
        Me.lstMoneyWon.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstMoneyWon.ForeColor = System.Drawing.Color.Black
        Me.lstMoneyWon.Items.AddRange(New Object() {"Question#    ", "1" & Global.Microsoft.VisualBasic.ChrW(9) & "$100", "2" & Global.Microsoft.VisualBasic.ChrW(9) & "$500", "3" & Global.Microsoft.VisualBasic.ChrW(9) & "$1000", "4" & Global.Microsoft.VisualBasic.ChrW(9) & "$2500", "5" & Global.Microsoft.VisualBasic.ChrW(9) & "$5000", "6" & Global.Microsoft.VisualBasic.ChrW(9) & "$10000", "7" & Global.Microsoft.VisualBasic.ChrW(9) & "$20000", "8" & Global.Microsoft.VisualBasic.ChrW(9) & "$40000", "9" & Global.Microsoft.VisualBasic.ChrW(9) & "$80000", "10" & Global.Microsoft.VisualBasic.ChrW(9) & "$100000", "11" & Global.Microsoft.VisualBasic.ChrW(9) & "$500000", "12" & Global.Microsoft.VisualBasic.ChrW(9) & "$1000000", "13" & Global.Microsoft.VisualBasic.ChrW(9) & "$5000000", "14" & Global.Microsoft.VisualBasic.ChrW(9) & "$10000000", "15" & Global.Microsoft.VisualBasic.ChrW(9) & "$15000000"})
        Me.lstMoneyWon.Location = New System.Drawing.Point(8, 32)
        Me.lstMoneyWon.Name = "lstMoneyWon"
        Me.lstMoneyWon.Size = New System.Drawing.Size(152, 212)
        Me.lstMoneyWon.TabIndex = 10
        Me.lstMoneyWon.Visible = False
        '
        'btnLifeLine
        '
        Me.btnLifeLine.BackColor = System.Drawing.SystemColors.Control
        Me.btnLifeLine.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnLifeLine.Location = New System.Drawing.Point(552, 104)
        Me.btnLifeLine.Name = "btnLifeLine"
        Me.btnLifeLine.Size = New System.Drawing.Size(104, 24)
        Me.btnLifeLine.TabIndex = 11
        Me.btnLifeLine.Text = "Life Line"
        Me.btnLifeLine.UseVisualStyleBackColor = False
        Me.btnLifeLine.Visible = False
        '
        'btnWalkAway
        '
        Me.btnWalkAway.BackColor = System.Drawing.SystemColors.Control
        Me.btnWalkAway.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnWalkAway.Location = New System.Drawing.Point(552, 144)
        Me.btnWalkAway.Name = "btnWalkAway"
        Me.btnWalkAway.Size = New System.Drawing.Size(104, 24)
        Me.btnWalkAway.TabIndex = 12
        Me.btnWalkAway.Text = "Walk Away"
        Me.btnWalkAway.UseVisualStyleBackColor = False
        Me.btnWalkAway.Visible = False
        '
        'lblQuestion
        '
        Me.lblQuestion.BackColor = System.Drawing.Color.Transparent
        Me.lblQuestion.Font = New System.Drawing.Font("Lucida Console", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion.Location = New System.Drawing.Point(8, 288)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.Size = New System.Drawing.Size(672, 88)
        Me.lblQuestion.TabIndex = 13
        Me.lblQuestion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblQuestion.Visible = False
        '
        'lblTimer
        '
        Me.lblTimer.AutoSize = True
        Me.lblTimer.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTimer.Location = New System.Drawing.Point(336, 240)
        Me.lblTimer.Name = "lblTimer"
        Me.lblTimer.Size = New System.Drawing.Size(0, 13)
        Me.lblTimer.TabIndex = 14
        Me.lblTimer.Visible = False
        '
        'tmrQuestionTimer
        '
        Me.tmrQuestionTimer.Interval = 1000
        '
        'btnStartNewGame
        '
        Me.btnStartNewGame.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStartNewGame.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btnStartNewGame.Location = New System.Drawing.Point(280, 280)
        Me.btnStartNewGame.Name = "btnStartNewGame"
        Me.btnStartNewGame.Size = New System.Drawing.Size(136, 64)
        Me.btnStartNewGame.TabIndex = 15
        Me.btnStartNewGame.Text = "Start New Game"
        Me.btnStartNewGame.UseVisualStyleBackColor = True
        '
        'prgTimeLeft
        '
        Me.prgTimeLeft.BackColor = System.Drawing.SystemColors.Control
        Me.prgTimeLeft.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.prgTimeLeft.Location = New System.Drawing.Point(64, 376)
        Me.prgTimeLeft.MarqueeAnimationSpeed = 0
        Me.prgTimeLeft.Name = "prgTimeLeft"
        Me.prgTimeLeft.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.prgTimeLeft.RightToLeftLayout = True
        Me.prgTimeLeft.Size = New System.Drawing.Size(288, 8)
        Me.prgTimeLeft.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.prgTimeLeft.TabIndex = 16
        Me.prgTimeLeft.UseWaitCursor = True
        Me.prgTimeLeft.Visible = False
        '
        'prgTimeRight
        '
        Me.prgTimeRight.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.prgTimeRight.Location = New System.Drawing.Point(352, 376)
        Me.prgTimeRight.MarqueeAnimationSpeed = 0
        Me.prgTimeRight.Name = "prgTimeRight"
        Me.prgTimeRight.Size = New System.Drawing.Size(304, 8)
        Me.prgTimeRight.TabIndex = 17
        Me.prgTimeRight.Visible = False
        '
        'tmrEndGameCelebration
        '
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Who_Wants_To_Be_A_Millionare.My.Resources.Resources.Wwtbam_uk_2010
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(695, 532)
        Me.Controls.Add(Me.prgTimeRight)
        Me.Controls.Add(Me.prgTimeLeft)
        Me.Controls.Add(Me.btnStartNewGame)
        Me.Controls.Add(Me.lblTimer)
        Me.Controls.Add(Me.lblQuestion)
        Me.Controls.Add(Me.btnWalkAway)
        Me.Controls.Add(Me.btnLifeLine)
        Me.Controls.Add(Me.lstMoneyWon)
        Me.Controls.Add(Me.picHintPicture)
        Me.Controls.Add(Me.lblOptionB)
        Me.Controls.Add(Me.lblOptionD)
        Me.Controls.Add(Me.lblOptionC)
        Me.Controls.Add(Me.lblOptionA)
        Me.Controls.Add(Me.btnAnswerChoiceC)
        Me.Controls.Add(Me.btnAnswerChoiceB)
        Me.Controls.Add(Me.btnAnswerChoiceD)
        Me.Controls.Add(Me.btnAnswerChoiceA)
        Me.Controls.Add(Me.mnuMenuStripOne)
        Me.DoubleBuffered = True
        Me.ForeColor = System.Drawing.Color.Yellow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.mnuMenuStripOne
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(711, 571)
        Me.Name = "Form1"
        Me.Text = "Who Wants To Be A Millionare?"
        Me.mnuMenuStripOne.ResumeLayout(False)
        Me.mnuMenuStripOne.PerformLayout()
        CType(Me.picHintPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnuMenuStripOne As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuProgram As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNewGame As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuWalkAway As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLifeLine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelpAndInstructions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAnswerChoiceA As System.Windows.Forms.Button
    Friend WithEvents btnAnswerChoiceD As System.Windows.Forms.Button
    Friend WithEvents btnAnswerChoiceB As System.Windows.Forms.Button
    Friend WithEvents btnAnswerChoiceC As System.Windows.Forms.Button
    Friend WithEvents lblOptionA As System.Windows.Forms.Label
    Friend WithEvents lblOptionC As System.Windows.Forms.Label
    Friend WithEvents lblOptionD As System.Windows.Forms.Label
    Friend WithEvents lblOptionB As System.Windows.Forms.Label
    Friend WithEvents picHintPicture As System.Windows.Forms.PictureBox
    Friend WithEvents lstMoneyWon As System.Windows.Forms.ListBox
    Friend WithEvents btnLifeLine As System.Windows.Forms.Button
    Friend WithEvents btnWalkAway As System.Windows.Forms.Button
    Friend WithEvents lblQuestion As System.Windows.Forms.Label
    Friend WithEvents lblTimer As System.Windows.Forms.Label
    Friend WithEvents tmrQuestionTimer As System.Windows.Forms.Timer
    Friend WithEvents btnStartNewGame As System.Windows.Forms.Button
    Friend WithEvents prgTimeLeft As System.Windows.Forms.ProgressBar
    Friend WithEvents prgTimeRight As System.Windows.Forms.ProgressBar
    Friend WithEvents tmrEndGameCelebration As System.Windows.Forms.Timer

End Class
