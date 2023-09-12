<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As DataGridViewCellStyle = New DataGridViewCellStyle()
        txtNumDice = New TextBox()
        txtDamage = New TextBox()
        txtFNPChance = New TextBox()
        txtSaveChance = New TextBox()
        txtWoundChance = New TextBox()
        txtHitChance = New TextBox()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        Label4 = New Label()
        Label5 = New Label()
        Label6 = New Label()
        Label7 = New Label()
        txtAutoWoundChance = New TextBox()
        Label8 = New Label()
        Label10 = New Label()
        Label11 = New Label()
        txtMWWoundChance = New TextBox()
        txtModifiedAPChance = New TextBox()
        Label13 = New Label()
        txtModifiedSaveChance = New TextBox()
        Label12 = New Label()
        rdbWoundsLost = New RadioButton()
        rdbCasualties = New RadioButton()
        dgvDistributionSpreadsheet = New DataGridView()
        Label16 = New Label()
        lblExpectedValue = New Label()
        lblStandardDeviation = New Label()
        Label19 = New Label()
        picDistributionHistogram = New PictureBox()
        btnCalculate = New Button()
        Label14 = New Label()
        Label17 = New Label()
        txtNumExtraHits = New TextBox()
        txtExtraHits = New TextBox()
        Label20 = New Label()
        Label21 = New Label()
        Label22 = New Label()
        Label23 = New Label()
        txtHitRerollChance = New TextBox()
        txtWoundRerollChance = New TextBox()
        txtSaveRerollChance = New TextBox()
        txtFNPRerollChance = New TextBox()
        txtRandomDamage = New TextBox()
        cmbRandomDamageType = New ComboBox()
        Label24 = New Label()
        Label25 = New Label()
        cmbRandomAttackType = New ComboBox()
        txtRandomAttacks = New TextBox()
        chkChooseBestAttacks = New CheckBox()
        Label26 = New Label()
        txtRerollAttacksChance = New TextBox()
        Label27 = New Label()
        txtRerollDamageChance = New TextBox()
        chkChooseBestDamage = New CheckBox()
        cmbTargetWounds = New ComboBox()
        cmbDiceType = New ComboBox()
        Label9 = New Label()
        chkRolloverDevWounds = New CheckBox()
        CType(dgvDistributionSpreadsheet, ComponentModel.ISupportInitialize).BeginInit()
        CType(picDistributionHistogram, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' txtNumDice
        ' 
        txtNumDice.Location = New Point(11, 70)
        txtNumDice.Margin = New Padding(2, 1, 2, 1)
        txtNumDice.Name = "txtNumDice"
        txtNumDice.Size = New Size(110, 23)
        txtNumDice.TabIndex = 0
        txtNumDice.Text = "10"
        ' 
        ' txtDamage
        ' 
        txtDamage.Location = New Point(571, 70)
        txtDamage.Margin = New Padding(2, 1, 2, 1)
        txtDamage.Name = "txtDamage"
        txtDamage.Size = New Size(110, 23)
        txtDamage.TabIndex = 17
        txtDamage.Text = "1"
        ' 
        ' txtFNPChance
        ' 
        txtFNPChance.Location = New Point(459, 70)
        txtFNPChance.Margin = New Padding(2, 1, 2, 1)
        txtFNPChance.Name = "txtFNPChance"
        txtFNPChance.Size = New Size(110, 23)
        txtFNPChance.TabIndex = 15
        txtFNPChance.Text = "1"
        ' 
        ' txtSaveChance
        ' 
        txtSaveChance.Location = New Point(348, 70)
        txtSaveChance.Margin = New Padding(2, 1, 2, 1)
        txtSaveChance.Name = "txtSaveChance"
        txtSaveChance.Size = New Size(110, 23)
        txtSaveChance.TabIndex = 12
        txtSaveChance.Text = "0.5"
        ' 
        ' txtWoundChance
        ' 
        txtWoundChance.Location = New Point(237, 70)
        txtWoundChance.Margin = New Padding(2, 1, 2, 1)
        txtWoundChance.Name = "txtWoundChance"
        txtWoundChance.Size = New Size(110, 23)
        txtWoundChance.TabIndex = 8
        txtWoundChance.Text = "0.5"
        ' 
        ' txtHitChance
        ' 
        txtHitChance.Location = New Point(126, 70)
        txtHitChance.Margin = New Padding(2, 1, 2, 1)
        txtHitChance.Name = "txtHitChance"
        txtHitChance.Size = New Size(110, 23)
        txtHitChance.TabIndex = 3
        txtHitChance.Text = "0.5"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(11, 39)
        Label1.Margin = New Padding(2, 0, 2, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(60, 30)
        Label1.TabIndex = 6
        Label1.Text = "Number " & vbCrLf & "of Attacks"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(126, 54)
        Label2.Margin = New Padding(2, 0, 2, 0)
        Label2.Name = "Label2"
        Label2.Size = New Size(66, 15)
        Label2.TabIndex = 7
        Label2.Text = "Hit Chance"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(237, 54)
        Label3.Margin = New Padding(2, 0, 2, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(89, 15)
        Label3.TabIndex = 8
        Label3.Text = "Wound Chance"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(348, 54)
        Label4.Margin = New Padding(2, 0, 2, 0)
        Label4.Name = "Label4"
        Label4.Size = New Size(101, 15)
        Label4.TabIndex = 9
        Label4.Text = "Save Chance (fail)"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(459, 54)
        Label5.Margin = New Padding(2, 0, 2, 0)
        Label5.Name = "Label5"
        Label5.Size = New Size(99, 15)
        Label5.TabIndex = 10
        Label5.Text = "FNP Chance (fail)"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(571, 54)
        Label6.Margin = New Padding(2, 0, 2, 0)
        Label6.Name = "Label6"
        Label6.Size = New Size(51, 15)
        Label6.TabIndex = 11
        Label6.Text = "Damage"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.ForeColor = SystemColors.GrayText
        Label7.Location = New Point(152, 9)
        Label7.Margin = New Padding(2, 0, 2, 0)
        Label7.Name = "Label7"
        Label7.Size = New Size(420, 30)
        Label7.TabIndex = 12
        Label7.Text = "Probabilities can be entered as fractions using / or as dice results using + and -" & vbCrLf & "Errors indicated by yellow textboxes. Double click to reset to default value."
        ' 
        ' txtAutoWoundChance
        ' 
        txtAutoWoundChance.Location = New Point(126, 168)
        txtAutoWoundChance.Margin = New Padding(2, 1, 2, 1)
        txtAutoWoundChance.Name = "txtAutoWoundChance"
        txtAutoWoundChance.Size = New Size(110, 23)
        txtAutoWoundChance.TabIndex = 5
        txtAutoWoundChance.Text = "0"
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Location = New Point(126, 137)
        Label8.Margin = New Padding(2, 0, 2, 0)
        Label8.Name = "Label8"
        Label8.Size = New Size(58, 30)
        Label8.TabIndex = 15
        Label8.Text = "Lethal Hit" & vbCrLf & "Chance"
        ' 
        ' Label10
        ' 
        Label10.AutoSize = True
        Label10.Location = New Point(237, 137)
        Label10.Margin = New Padding(2, 0, 2, 0)
        Label10.Name = "Label10"
        Label10.Size = New Size(111, 30)
        Label10.TabIndex = 20
        Label10.Text = "Devastating Wound" & vbCrLf & "Chance"
        ' 
        ' Label11
        ' 
        Label11.AutoSize = True
        Label11.Location = New Point(237, 193)
        Label11.Margin = New Padding(2, 0, 2, 0)
        Label11.Name = "Label11"
        Label11.Size = New Size(73, 30)
        Label11.TabIndex = 19
        Label11.Text = "Modified AP" & vbCrLf & "Chance"
        ' 
        ' txtMWWoundChance
        ' 
        txtMWWoundChance.Location = New Point(237, 168)
        txtMWWoundChance.Margin = New Padding(2, 1, 2, 1)
        txtMWWoundChance.Name = "txtMWWoundChance"
        txtMWWoundChance.Size = New Size(110, 23)
        txtMWWoundChance.TabIndex = 10
        txtMWWoundChance.Text = "0"
        ' 
        ' txtModifiedAPChance
        ' 
        txtModifiedAPChance.Location = New Point(237, 224)
        txtModifiedAPChance.Margin = New Padding(2, 1, 2, 1)
        txtModifiedAPChance.Name = "txtModifiedAPChance"
        txtModifiedAPChance.Size = New Size(110, 23)
        txtModifiedAPChance.TabIndex = 11
        txtModifiedAPChance.Text = "0"
        ' 
        ' Label13
        ' 
        Label13.AutoSize = True
        Label13.Location = New Point(348, 193)
        Label13.Margin = New Padding(2, 0, 2, 0)
        Label13.Name = "Label13"
        Label13.Size = New Size(82, 30)
        Label13.TabIndex = 23
        Label13.Text = "Modified Save" & vbCrLf & "Chance (fail)"
        ' 
        ' txtModifiedSaveChance
        ' 
        txtModifiedSaveChance.Location = New Point(348, 224)
        txtModifiedSaveChance.Margin = New Padding(2, 1, 2, 1)
        txtModifiedSaveChance.Name = "txtModifiedSaveChance"
        txtModifiedSaveChance.Size = New Size(110, 23)
        txtModifiedSaveChance.TabIndex = 14
        txtModifiedSaveChance.Text = "0"
        ' 
        ' Label12
        ' 
        Label12.AutoSize = True
        Label12.Location = New Point(570, 206)
        Label12.Margin = New Padding(2, 0, 2, 0)
        Label12.Name = "Label12"
        Label12.Size = New Size(90, 15)
        Label12.TabIndex = 25
        Label12.Text = "Wounds/model"
        ' 
        ' rdbWoundsLost
        ' 
        rdbWoundsLost.AutoSize = True
        rdbWoundsLost.Checked = True
        rdbWoundsLost.Location = New Point(502, 263)
        rdbWoundsLost.Margin = New Padding(2, 1, 2, 1)
        rdbWoundsLost.Name = "rdbWoundsLost"
        rdbWoundsLost.Size = New Size(94, 19)
        rdbWoundsLost.TabIndex = 21
        rdbWoundsLost.TabStop = True
        rdbWoundsLost.Text = "Wounds Lost"
        rdbWoundsLost.UseVisualStyleBackColor = True
        ' 
        ' rdbCasualties
        ' 
        rdbCasualties.AutoSize = True
        rdbCasualties.Location = New Point(600, 263)
        rdbCasualties.Margin = New Padding(2, 1, 2, 1)
        rdbCasualties.Name = "rdbCasualties"
        rdbCasualties.Size = New Size(78, 19)
        rdbCasualties.TabIndex = 22
        rdbCasualties.Text = "Casualties"
        rdbCasualties.UseVisualStyleBackColor = True
        ' 
        ' dgvDistributionSpreadsheet
        ' 
        dgvDistributionSpreadsheet.AllowUserToAddRows = False
        dgvDistributionSpreadsheet.AllowUserToDeleteRows = False
        dgvDistributionSpreadsheet.ColumnHeadersHeight = 46
        dgvDistributionSpreadsheet.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = SystemColors.Window
        DataGridViewCellStyle1.Font = New Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point)
        DataGridViewCellStyle1.ForeColor = SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = DataGridViewTriState.True
        dgvDistributionSpreadsheet.DefaultCellStyle = DataGridViewCellStyle1
        dgvDistributionSpreadsheet.Location = New Point(704, 54)
        dgvDistributionSpreadsheet.Margin = New Padding(2, 1, 2, 1)
        dgvDistributionSpreadsheet.Name = "dgvDistributionSpreadsheet"
        dgvDistributionSpreadsheet.ReadOnly = True
        dgvDistributionSpreadsheet.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dgvDistributionSpreadsheet.RowHeadersVisible = False
        dgvDistributionSpreadsheet.RowHeadersWidth = 82
        dgvDistributionSpreadsheet.RowTemplate.Height = 20
        dgvDistributionSpreadsheet.RowTemplate.Resizable = DataGridViewTriState.False
        dgvDistributionSpreadsheet.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvDistributionSpreadsheet.Size = New Size(240, 468)
        dgvDistributionSpreadsheet.TabIndex = 24
        ' 
        ' Label16
        ' 
        Label16.AutoSize = True
        Label16.Location = New Point(600, 22)
        Label16.Margin = New Padding(2, 0, 2, 0)
        Label16.Name = "Label16"
        Label16.Size = New Size(86, 15)
        Label16.TabIndex = 31
        Label16.Text = "Expected Value"
        ' 
        ' lblExpectedValue
        ' 
        lblExpectedValue.BorderStyle = BorderStyle.FixedSingle
        lblExpectedValue.Location = New Point(690, 21)
        lblExpectedValue.Margin = New Padding(2, 0, 2, 0)
        lblExpectedValue.Name = "lblExpectedValue"
        lblExpectedValue.Size = New Size(76, 21)
        lblExpectedValue.TabIndex = 32
        ' 
        ' lblStandardDeviation
        ' 
        lblStandardDeviation.BorderStyle = BorderStyle.FixedSingle
        lblStandardDeviation.Location = New Point(842, 18)
        lblStandardDeviation.Margin = New Padding(2, 0, 2, 0)
        lblStandardDeviation.Name = "lblStandardDeviation"
        lblStandardDeviation.Size = New Size(102, 21)
        lblStandardDeviation.TabIndex = 34
        ' 
        ' Label19
        ' 
        Label19.AutoSize = True
        Label19.Location = New Point(781, 12)
        Label19.Margin = New Padding(2, 0, 2, 0)
        Label19.Name = "Label19"
        Label19.Size = New Size(57, 30)
        Label19.TabIndex = 33
        Label19.Text = "Standard" & vbCrLf & "Deviation"
        ' 
        ' picDistributionHistogram
        ' 
        picDistributionHistogram.BackColor = Color.White
        picDistributionHistogram.BorderStyle = BorderStyle.FixedSingle
        picDistributionHistogram.Location = New Point(33, 301)
        picDistributionHistogram.Margin = New Padding(2, 1, 2, 1)
        picDistributionHistogram.Name = "picDistributionHistogram"
        picDistributionHistogram.Size = New Size(667, 221)
        picDistributionHistogram.TabIndex = 35
        picDistributionHistogram.TabStop = False
        ' 
        ' btnCalculate
        ' 
        btnCalculate.Location = New Point(11, 235)
        btnCalculate.Margin = New Padding(2, 1, 2, 1)
        btnCalculate.Name = "btnCalculate"
        btnCalculate.Size = New Size(102, 31)
        btnCalculate.TabIndex = 36
        btnCalculate.Text = "Calculate"
        btnCalculate.UseVisualStyleBackColor = True
        ' 
        ' Label14
        ' 
        Label14.AutoSize = True
        Label14.Location = New Point(126, 248)
        Label14.Margin = New Padding(2, 0, 2, 0)
        Label14.Name = "Label14"
        Label14.Size = New Size(67, 15)
        Label14.TabIndex = 40
        Label14.Text = "# Extra Hits"
        ' 
        ' Label17
        ' 
        Label17.AutoSize = True
        Label17.Location = New Point(126, 193)
        Label17.Margin = New Padding(2, 0, 2, 0)
        Label17.Name = "Label17"
        Label17.Size = New Size(82, 30)
        Label17.TabIndex = 39
        Label17.Text = "Sustained Hits" & vbCrLf & "Chance"
        ' 
        ' txtNumExtraHits
        ' 
        txtNumExtraHits.Location = New Point(126, 263)
        txtNumExtraHits.Margin = New Padding(2, 1, 2, 1)
        txtNumExtraHits.Name = "txtNumExtraHits"
        txtNumExtraHits.Size = New Size(110, 23)
        txtNumExtraHits.TabIndex = 7
        txtNumExtraHits.Text = "1"
        ' 
        ' txtExtraHits
        ' 
        txtExtraHits.Location = New Point(126, 224)
        txtExtraHits.Margin = New Padding(2, 1, 2, 1)
        txtExtraHits.Name = "txtExtraHits"
        txtExtraHits.Size = New Size(110, 23)
        txtExtraHits.TabIndex = 6
        txtExtraHits.Text = "0"
        ' 
        ' Label20
        ' 
        Label20.AutoSize = True
        Label20.Location = New Point(459, 97)
        Label20.Margin = New Padding(2, 0, 2, 0)
        Label20.Name = "Label20"
        Label20.Size = New Size(80, 15)
        Label20.TabIndex = 51
        Label20.Text = "Reroll Chance"
        ' 
        ' Label21
        ' 
        Label21.AutoSize = True
        Label21.Location = New Point(348, 97)
        Label21.Margin = New Padding(2, 0, 2, 0)
        Label21.Name = "Label21"
        Label21.Size = New Size(80, 15)
        Label21.TabIndex = 50
        Label21.Text = "Reroll Chance"
        ' 
        ' Label22
        ' 
        Label22.AutoSize = True
        Label22.Location = New Point(237, 97)
        Label22.Margin = New Padding(2, 0, 2, 0)
        Label22.Name = "Label22"
        Label22.Size = New Size(80, 15)
        Label22.TabIndex = 49
        Label22.Text = "Reroll Chance"
        ' 
        ' Label23
        ' 
        Label23.AutoSize = True
        Label23.Location = New Point(126, 97)
        Label23.Margin = New Padding(2, 0, 2, 0)
        Label23.Name = "Label23"
        Label23.Size = New Size(80, 15)
        Label23.TabIndex = 48
        Label23.Text = "Reroll Chance"
        ' 
        ' txtHitRerollChance
        ' 
        txtHitRerollChance.Location = New Point(126, 113)
        txtHitRerollChance.Margin = New Padding(2, 1, 2, 1)
        txtHitRerollChance.Name = "txtHitRerollChance"
        txtHitRerollChance.Size = New Size(110, 23)
        txtHitRerollChance.TabIndex = 4
        txtHitRerollChance.Text = "0"
        ' 
        ' txtWoundRerollChance
        ' 
        txtWoundRerollChance.Location = New Point(237, 113)
        txtWoundRerollChance.Margin = New Padding(2, 1, 2, 1)
        txtWoundRerollChance.Name = "txtWoundRerollChance"
        txtWoundRerollChance.Size = New Size(110, 23)
        txtWoundRerollChance.TabIndex = 9
        txtWoundRerollChance.Text = "0"
        ' 
        ' txtSaveRerollChance
        ' 
        txtSaveRerollChance.Location = New Point(348, 113)
        txtSaveRerollChance.Margin = New Padding(2, 1, 2, 1)
        txtSaveRerollChance.Name = "txtSaveRerollChance"
        txtSaveRerollChance.Size = New Size(110, 23)
        txtSaveRerollChance.TabIndex = 13
        txtSaveRerollChance.Text = "0"
        ' 
        ' txtFNPRerollChance
        ' 
        txtFNPRerollChance.Location = New Point(459, 113)
        txtFNPRerollChance.Margin = New Padding(2, 1, 2, 1)
        txtFNPRerollChance.Name = "txtFNPRerollChance"
        txtFNPRerollChance.Size = New Size(110, 23)
        txtFNPRerollChance.TabIndex = 16
        txtFNPRerollChance.Text = "0"
        ' 
        ' txtRandomDamage
        ' 
        txtRandomDamage.Location = New Point(579, 113)
        txtRandomDamage.Margin = New Padding(2, 1, 2, 1)
        txtRandomDamage.Name = "txtRandomDamage"
        txtRandomDamage.Size = New Size(42, 23)
        txtRandomDamage.TabIndex = 18
        txtRandomDamage.Text = "0"
        txtRandomDamage.TextAlign = HorizontalAlignment.Right
        ' 
        ' cmbRandomDamageType
        ' 
        cmbRandomDamageType.DropDownStyle = ComboBoxStyle.DropDownList
        cmbRandomDamageType.FormattingEnabled = True
        cmbRandomDamageType.Items.AddRange(New Object() {"D3", "D6"})
        cmbRandomDamageType.Location = New Point(626, 113)
        cmbRandomDamageType.Name = "cmbRandomDamageType"
        cmbRandomDamageType.Size = New Size(52, 23)
        cmbRandomDamageType.TabIndex = 19
        ' 
        ' Label24
        ' 
        Label24.AutoSize = True
        Label24.Location = New Point(570, 94)
        Label24.Margin = New Padding(2, 0, 2, 0)
        Label24.Name = "Label24"
        Label24.Size = New Size(110, 15)
        Label24.TabIndex = 54
        Label24.Text = "+ Random Damage"
        ' 
        ' Label25
        ' 
        Label25.AutoSize = True
        Label25.Location = New Point(12, 97)
        Label25.Margin = New Padding(2, 0, 2, 0)
        Label25.Name = "Label25"
        Label25.Size = New Size(105, 15)
        Label25.TabIndex = 57
        Label25.Text = "+ Random Attacks"
        ' 
        ' cmbRandomAttackType
        ' 
        cmbRandomAttackType.DropDownStyle = ComboBoxStyle.DropDownList
        cmbRandomAttackType.FormattingEnabled = True
        cmbRandomAttackType.Items.AddRange(New Object() {"D3", "D6"})
        cmbRandomAttackType.Location = New Point(66, 113)
        cmbRandomAttackType.Name = "cmbRandomAttackType"
        cmbRandomAttackType.Size = New Size(52, 23)
        cmbRandomAttackType.TabIndex = 2
        ' 
        ' txtRandomAttacks
        ' 
        txtRandomAttacks.Location = New Point(22, 113)
        txtRandomAttacks.Margin = New Padding(2, 1, 2, 1)
        txtRandomAttacks.Name = "txtRandomAttacks"
        txtRandomAttacks.Size = New Size(42, 23)
        txtRandomAttacks.TabIndex = 1
        txtRandomAttacks.Text = "0"
        txtRandomAttacks.TextAlign = HorizontalAlignment.Right
        ' 
        ' chkChooseBestAttacks
        ' 
        chkChooseBestAttacks.CheckAlign = ContentAlignment.BottomCenter
        chkChooseBestAttacks.Enabled = False
        chkChooseBestAttacks.Location = New Point(69, 140)
        chkChooseBestAttacks.Name = "chkChooseBestAttacks"
        chkChooseBestAttacks.Size = New Size(52, 52)
        chkChooseBestAttacks.TabIndex = 62
        chkChooseBestAttacks.Text = "Choose" & vbCrLf & "Best?"
        chkChooseBestAttacks.TextAlign = ContentAlignment.TopCenter
        chkChooseBestAttacks.UseVisualStyleBackColor = True
        ' 
        ' Label26
        ' 
        Label26.AutoSize = True
        Label26.Location = New Point(11, 137)
        Label26.Margin = New Padding(2, 0, 2, 0)
        Label26.Name = "Label26"
        Label26.Size = New Size(47, 30)
        Label26.TabIndex = 61
        Label26.Text = "Reroll" & vbCrLf & "Chance"
        ' 
        ' txtRerollAttacksChance
        ' 
        txtRerollAttacksChance.Enabled = False
        txtRerollAttacksChance.Location = New Point(12, 172)
        txtRerollAttacksChance.Margin = New Padding(2, 1, 2, 1)
        txtRerollAttacksChance.Name = "txtRerollAttacksChance"
        txtRerollAttacksChance.Size = New Size(52, 23)
        txtRerollAttacksChance.TabIndex = 60
        txtRerollAttacksChance.Text = "0"
        ' 
        ' Label27
        ' 
        Label27.AutoSize = True
        Label27.Location = New Point(571, 140)
        Label27.Margin = New Padding(2, 0, 2, 0)
        Label27.Name = "Label27"
        Label27.Size = New Size(37, 15)
        Label27.TabIndex = 64
        Label27.Text = "Reroll"
        ' 
        ' txtRerollDamageChance
        ' 
        txtRerollDamageChance.Enabled = False
        txtRerollDamageChance.Location = New Point(571, 156)
        txtRerollDamageChance.Margin = New Padding(2, 1, 2, 1)
        txtRerollDamageChance.Name = "txtRerollDamageChance"
        txtRerollDamageChance.Size = New Size(63, 23)
        txtRerollDamageChance.TabIndex = 63
        txtRerollDamageChance.Text = "0"
        ' 
        ' chkChooseBestDamage
        ' 
        chkChooseBestDamage.CheckAlign = ContentAlignment.BottomCenter
        chkChooseBestDamage.Enabled = False
        chkChooseBestDamage.Location = New Point(629, 140)
        chkChooseBestDamage.Name = "chkChooseBestDamage"
        chkChooseBestDamage.Size = New Size(52, 39)
        chkChooseBestDamage.TabIndex = 64
        chkChooseBestDamage.Text = "Best?"
        chkChooseBestDamage.TextAlign = ContentAlignment.TopCenter
        chkChooseBestDamage.UseVisualStyleBackColor = True
        ' 
        ' cmbTargetWounds
        ' 
        cmbTargetWounds.DropDownStyle = ComboBoxStyle.DropDownList
        cmbTargetWounds.FormattingEnabled = True
        cmbTargetWounds.Items.AddRange(New Object() {"Single Target", "1", "2", "3", "4", "5", "6"})
        cmbTargetWounds.Location = New Point(561, 224)
        cmbTargetWounds.Name = "cmbTargetWounds"
        cmbTargetWounds.Size = New Size(117, 23)
        cmbTargetWounds.TabIndex = 20
        ' 
        ' cmbDiceType
        ' 
        cmbDiceType.DropDownStyle = ComboBoxStyle.DropDownList
        cmbDiceType.FormattingEnabled = True
        cmbDiceType.Items.AddRange(New Object() {"D3", "D6", "D8", "D10", "D12", "D20"})
        cmbDiceType.Location = New Point(89, 9)
        cmbDiceType.Name = "cmbDiceType"
        cmbDiceType.Size = New Size(58, 23)
        cmbDiceType.TabIndex = 66
        ' 
        ' Label9
        ' 
        Label9.AutoSize = True
        Label9.Location = New Point(26, 12)
        Label9.Name = "Label9"
        Label9.Size = New Size(57, 15)
        Label9.TabIndex = 66
        Label9.Text = "Dice Type"
        ' 
        ' chkRolloverDevWounds
        ' 
        chkRolloverDevWounds.AutoSize = True
        chkRolloverDevWounds.Location = New Point(352, 170)
        chkRolloverDevWounds.Name = "chkRolloverDevWounds"
        chkRolloverDevWounds.Size = New Size(186, 19)
        chkRolloverDevWounds.TabIndex = 65
        chkRolloverDevWounds.Text = "Rollover Devastating Wounds?"
        chkRolloverDevWounds.UseVisualStyleBackColor = True
        ' 
        ' frmMain
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(955, 533)
        Controls.Add(chkRolloverDevWounds)
        Controls.Add(Label9)
        Controls.Add(cmbDiceType)
        Controls.Add(cmbTargetWounds)
        Controls.Add(txtRerollDamageChance)
        Controls.Add(txtRerollAttacksChance)
        Controls.Add(cmbRandomAttackType)
        Controls.Add(txtRandomAttacks)
        Controls.Add(cmbRandomDamageType)
        Controls.Add(txtRandomDamage)
        Controls.Add(txtHitRerollChance)
        Controls.Add(txtWoundRerollChance)
        Controls.Add(txtSaveRerollChance)
        Controls.Add(txtFNPRerollChance)
        Controls.Add(txtNumExtraHits)
        Controls.Add(txtExtraHits)
        Controls.Add(btnCalculate)
        Controls.Add(lblStandardDeviation)
        Controls.Add(lblExpectedValue)
        Controls.Add(dgvDistributionSpreadsheet)
        Controls.Add(rdbCasualties)
        Controls.Add(rdbWoundsLost)
        Controls.Add(txtModifiedSaveChance)
        Controls.Add(txtMWWoundChance)
        Controls.Add(txtModifiedAPChance)
        Controls.Add(txtAutoWoundChance)
        Controls.Add(Label7)
        Controls.Add(txtHitChance)
        Controls.Add(txtWoundChance)
        Controls.Add(txtSaveChance)
        Controls.Add(txtFNPChance)
        Controls.Add(txtDamage)
        Controls.Add(txtNumDice)
        Controls.Add(picDistributionHistogram)
        Controls.Add(Label1)
        Controls.Add(Label2)
        Controls.Add(Label3)
        Controls.Add(Label16)
        Controls.Add(Label19)
        Controls.Add(Label25)
        Controls.Add(Label23)
        Controls.Add(Label22)
        Controls.Add(Label26)
        Controls.Add(chkChooseBestAttacks)
        Controls.Add(Label8)
        Controls.Add(Label17)
        Controls.Add(Label14)
        Controls.Add(Label11)
        Controls.Add(Label10)
        Controls.Add(Label13)
        Controls.Add(Label12)
        Controls.Add(Label4)
        Controls.Add(Label21)
        Controls.Add(Label5)
        Controls.Add(Label20)
        Controls.Add(Label6)
        Controls.Add(Label24)
        Controls.Add(Label27)
        Controls.Add(chkChooseBestDamage)
        FormBorderStyle = FormBorderStyle.FixedSingle
        Margin = New Padding(2, 1, 2, 1)
        MaximizeBox = False
        MaximumSize = New Size(971, 572)
        MinimumSize = New Size(971, 572)
        Name = "frmMain"
        StartPosition = FormStartPosition.CenterScreen
        Text = "The 40Kalculator"
        CType(dgvDistributionSpreadsheet, ComponentModel.ISupportInitialize).EndInit()
        CType(picDistributionHistogram, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtNumDice As TextBox
    Friend WithEvents txtDamage As TextBox
    Friend WithEvents txtFNPChance As TextBox
    Friend WithEvents txtSaveChance As TextBox
    Friend WithEvents txtWoundChance As TextBox
    Friend WithEvents txtHitChance As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents txtAutoWoundChance As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents txtMWWoundChance As TextBox
    Friend WithEvents txtModifiedAPChance As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txtModifiedSaveChance As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents rdbWoundsLost As RadioButton
    Friend WithEvents rdbCasualties As RadioButton
    Friend WithEvents dgvDistributionSpreadsheet As DataGridView
    Friend WithEvents Label16 As Label
    Friend WithEvents lblExpectedValue As Label
    Friend WithEvents lblStandardDeviation As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents picDistributionHistogram As PictureBox
    Friend WithEvents btnCalculate As Button
    Friend WithEvents Label14 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents txtNumExtraHits As TextBox
    Friend WithEvents txtExtraHits As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents Label23 As Label
    Friend WithEvents txtHitRerollChance As TextBox
    Friend WithEvents txtWoundRerollChance As TextBox
    Friend WithEvents txtSaveRerollChance As TextBox
    Friend WithEvents txtFNPRerollChance As TextBox
    Friend WithEvents txtRandomDamage As TextBox
    Friend WithEvents cmbRandomDamageType As ComboBox
    Friend WithEvents cmbRandomAttackType As ComboBox
    Friend WithEvents Label24 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents txtRandomAttacks As TextBox
    Friend WithEvents chkChooseBestAttacks As CheckBox
    Friend WithEvents Label26 As Label
    Friend WithEvents txtRerollAttacksChance As TextBox
    Friend WithEvents Label27 As Label
    Friend WithEvents txtRerollDamageChance As TextBox
    Friend WithEvents chkChooseBestDamage As CheckBox
    Friend WithEvents cmbTargetWounds As ComboBox
    Friend WithEvents cmbDiceType As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents chkRolloverDevWounds As CheckBox
End Class
