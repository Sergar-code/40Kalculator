Imports System.Reflection.Metadata.Ecma335

Public Class frmMain

    'The following are stored to quickly switch between damage/casualties data
    'Stores the PDF of casualties caused. This ignores excess wounds and partially wounded models
    Dim CasualtiesPDF() As Double

    'Stores the PDF of direct damage caused. This ignores the target health and just calculates total damage dealt.
    Dim DamagePDF() As Double

    'Number of wounds per model is stored to give meaningful labels to the damage column
    Dim Wounds As Integer

    Dim ScaleX As Decimal
    Dim ScaleY As Decimal

    'Define a list of collumns as displayed in the graph screen
    Dim ColumnArray As List(Of Rectangle)

    'Define a list to hold the columns that are currently selected
    Dim SelectedColumns As List(Of Rectangle) = New List(Of Rectangle)

    'Track whether the user is selecting on dgvDistributionSpreadsheet or picDistributionHistogram
    Dim HistogramSelection As Boolean = False
    Dim SpreadsheetSelection As Boolean = False

    'Used to convert the different dice type indices from cmbDiceType into the number of faces
    Enum DiceType
        D3 = 0
        D6 = 1
        D8 = 2
        D10 = 3
        D12 = 4
        D20 = 5
    End Enum

    Private Function GetDiceSize(DType As DiceType) As Integer
        Select Case DType
            Case DiceType.D3
                Return 3
            Case DiceType.D6
                Return 6
            Case DiceType.D8
                Return 8
            Case DiceType.D10
                Return 10
            Case DiceType.D12
                Return 12
            Case DiceType.D20
                Return 20
        End Select
        Return 1
    End Function

    Private Function TextSize(ByVal HeaderText As String, ByVal HeaderFont As Font, ByVal WrapLine As Boolean) As Size
        Dim HeaderSize As Size

        If WrapLine Then
            Dim EachLine() As String = HeaderText.Split(" ")
            Dim MaxWidth As Integer
            Dim HeaderHeight As Integer
            For Each HeaderLine As String In EachLine
                MaxWidth = Math.Max(TextRenderer.MeasureText(HeaderLine, HeaderFont).Width, MaxWidth)
                HeaderHeight += TextRenderer.MeasureText(HeaderLine, HeaderFont).Height + 2
            Next
            HeaderSize.Width = MaxWidth
            HeaderSize.Height = HeaderHeight
        Else
            HeaderSize = TextRenderer.MeasureText(HeaderText, HeaderFont)
        End If
        Return HeaderSize
    End Function

    'Returns the distribution, as it is symmetrical this is faster by removing redundant calculations.
    'Uses the logarithm approach to reduce factorial load on the floating point variables

    Private Function BinomPDF(NumTrials As Integer, SuccessChance As Double, CurrentStep As String) As Double()

        If SuccessChance < 0 Or SuccessChance > 1 Then
            If Math.Round(SuccessChance, 5) = 0 Or Math.Round(SuccessChance, 5) = 1 Then
                SuccessChance = Math.Round(SuccessChance, 5)
            Else
                MsgBox("Parameter Error in BinomPDF at step: " + CurrentStep, vbOKOnly, "Error")
                Dim EmptyDist() As Double = New Double() {}
                Return EmptyDist
            End If


        End If

        'Trivial Case of 0 success on 0 trials
        If NumTrials = 0 Then
            Dim ZeroDist(0) As Double
            ZeroDist(0) = 1
            Return ZeroDist
        End If

        Dim OutputDist(NumTrials) As Double
        'Only procceed to midpoint, as combination distribution is symmetrical
        For NumSuccesses = 0 To Math.Ceiling(NumTrials / 2)
            'Finds the larger and smaller of r and (n-r). Used to optimise factorial calcualtion.
            'Larger is stored in Divisors(1). It doesn't matter if they are equal.
            Dim Divisors(1) As Integer
            Divisors(0) = NumSuccesses
            Divisors(1) = NumTrials - NumSuccesses
            If Divisors(1) < Divisors(0) Then
                Dim PlaceHolder As Integer
                PlaceHolder = Divisors(0)
                Divisors(0) = Divisors(1)
                Divisors(1) = PlaceHolder
            End If
            Dim Combinations As Double = 0
            'All of the divisor factorials will be smaller than the numerator factorials, so these are calculated first to preserve precision
            For Iterator = 2 To Divisors(0)
                Combinations -= Math.Log(Iterator)
            Next
            For Iterator = Divisors(1) + 1 To NumTrials
                Combinations += Math.Log(Iterator)
            Next

            OutputDist(NumSuccesses) = Math.Exp(Combinations) * (SuccessChance ^ NumSuccesses) * ((1 - SuccessChance) ^ (NumTrials - NumSuccesses))
            OutputDist(NumTrials - NumSuccesses) = Math.Exp(Combinations) * (SuccessChance ^ (NumTrials - NumSuccesses)) * ((1 - SuccessChance) ^ NumSuccesses)
        Next

        Return OutputDist

    End Function

    Private Function DiceSumCombinations(PrevPDF As List(Of ULong), DiceSize As Integer) As List(Of ULong)

        'The function returns an array of integers to preserve exactness.
        'Probabilities are calculated by dividing each entry by the sum of the entries.

        Dim OutputPDF As List(Of ULong) = New List(Of ULong)

        'The PDF for 0 dice is 100% probability of a sum = 0, so this is returned on receiving an empty array
        If PrevPDF.Count < 1 Then
            OutputPDF.Add(1)
            Return OutputPDF
        End If

        Dim RunningSum As List(Of ULong) = New List(Of ULong)
        RunningSum.Add(0)
        OutputPDF.Add(0)

        'This algorithm is based on the interpretation of dice sum as diagonal cross-sections of an n-dimensional hypercube
        'where n is the number of dice.
        For index = 1 To PrevPDF.Count + DiceSize - 1
            'the next dimension of diagonal cross-sections is determined by summing the previous dimensional cross-sections
            'For example, with 3 dice this resembles the triangle numbers
            If index <= PrevPDF.Count Then
                RunningSum.Add(RunningSum(index - 1) + PrevPDF(index - 1))
            Else
                'It is assumed that each pdf is surrounded by an infinite sequence of 0's
                'So when the previous PDF is exhausted, the sum is omitted as it would be adding 0
                RunningSum.Add(RunningSum(index - 1))
            End If
            If index - DiceSize >= 0 Then
                'The cross-section has a maximum length of DiceSize (typically 6 for a 6-sided dice)
                'This subtracts the extraneous rows of the cross-section
                OutputPDF.Add(RunningSum(index) - RunningSum(index - DiceSize))
            Else
                OutputPDF.Add(RunningSum(index))
            End If
        Next

        Return OutputPDF

    End Function

    'Returns a PDF of the dicesum including a fixed value. Accepts DiceSize as the number of faces on the dice
    Private Function DiceSumPDF(FixedNum As Integer, DiceNum As Integer, DiceSize As Integer) As Double()

        'Represents a 100% probability of a 0 sum result from 0 dice
        Dim RunningCombinations As List(Of ULong) = New List(Of ULong)
        RunningCombinations.Add(1)

        'Iterate through to get the combinations for the given number of dice
        For DiceCount = 1 To DiceNum
            RunningCombinations = DiceSumCombinations(RunningCombinations, DiceSize)
        Next

        Dim CombinationSum As ULong = 0
        For Each Combination As ULong In RunningCombinations
            CombinationSum += Combination
        Next

        Dim OutputPDF(FixedNum + RunningCombinations.Count - 1) As Double

        For RandomNum = 0 To RunningCombinations.Count - 1
            OutputPDF(FixedNum + RandomNum) = RunningCombinations(RandomNum) / CombinationSum
        Next

        Return OutputPDF

    End Function

    Private Function SumDistributions(ByVal Dist1() As Double, ByVal Dist2() As Double, MaxSize As Integer) As Double()

        ' If either input distribution is empty, then return the other distribution unchanged
        'If both distributions are empty this will return an empty distribution, which is the correct behaviour
        If Dist1.Count = 0 Then
            Return Dist2
        End If
        If Dist2.Count = 0 Then
            Return Dist1
        End If

        'Input distributions start at 0, requiring subtracting 1 from the count for each. The maximum possible size is the sum of 
        'the largest indices of each distribution, representing the maximum possible result from each
        Dim OutputDist(Math.Min(Dist1.Count + Dist2.Count - 2, MaxSize)) As Double

        'Probabilities are referenced with indices i1 (Dist1), i2 = i3 - i1 (Dist2 - not used) and i3 (OutputDist)
        For i3 = 0 To OutputDist.Count - 1
            Dim CurrentProbability As Double = 0
            'For each possible sum in the output distribution...
            For i1 = 0 To i3
                'Find each combination of values that would sum to this output sum and...
                If Dist1.Count > i1 And Dist2.Count > i3 - i1 Then
                    'Calculate the probability of this combination by multiplying the distribution probabilities
                    CurrentProbability += Dist1(i1) * Dist2(i3 - i1)
                End If
            Next
            'If Double.IsNaN(CurrentProbability) Or Double.IsInfinity(CurrentProbability) Then MsgBox("Invalid probability generated in SumDistributions Function", vbOKOnly, "Error")

            If CurrentProbability <= 1 Then
                OutputDist(i3) = CurrentProbability
            End If
        Next

        Return OutputDist

    End Function

    'Calculates chance of a regular success (no crit) including rerolls.
    'Accounts for the possibility of rerolling successes (such as fishing for crits)
    Private Function CalculateHitChance(HitChance As Double, CritChance As Double, RerollChance As Double) As Double
        Dim CorrectedRerollChance As Double = Math.Min(1 - CritChance, RerollChance)
        Return Math.Min(HitChance - CritChance, 1 - CorrectedRerollChance - CritChance) + CorrectedRerollChance * (HitChance - CritChance)
    End Function

    'Calculates chance of critical success including rerolls
    'Accounts for the possibility of rerolling crits (Probably an error, but following GIGO principle here)
    'If the code is ever modified for different chances for different crits, this could be useful.
    Private Function CalculateCritChance(CritChance As Double, RerollChance As Double) As Double
        Dim CorrectedRerollChance As Double = Math.Min(1 - CritChance, RerollChance)
        Return Math.Min(CritChance, 1 - CorrectedRerollChance) + CorrectedRerollChance * CritChance
    End Function

    Private Function ParseProbability(ByVal Text As String, ByVal DiceFaces As Integer) As Decimal
        If Not (Text.Contains("/")) And Not (Text.Contains("+")) And Not (Text.Contains("-")) Then
            If IsNumeric(Text) AndAlso Decimal.Parse(Text) <= 1 AndAlso Decimal.Parse(Text) >= 0 Then
                Return Decimal.Parse(Text)
            Else
                Return -1
            End If
        ElseIf Text.Contains("/") And Not (Text.Contains("+") Or Text.Contains("-")) Then
            Return ParseFraction(Text)
        ElseIf Not (Text.Contains("/")) Then
            Return ParseDiceChance(Text, DiceFaces)
        Else
            Return -1
        End If
    End Function

    Private Function ParseFraction(ByVal Text As String) As Decimal
        'Converts fraction entered as text using "/" to a decimal. Returns "-1" on error.
        '-1 is detected as error since probabilities must be between 0 and 1. Not suitable for other use cases
        'Immediately returns error if more than one "/" detected
        Dim Fraction() As String = Split(Text, "/")
        If Fraction.Count = 2 And IsNumeric(Fraction(0)) And IsNumeric(Fraction(1)) Then
            If Format(Fraction(0)) / Format(Fraction(1)) >= 0 And Format(Fraction(0)) / Format(Fraction(1)) <= 1 Then
                Return Format(Fraction(0)) / Format(Fraction(1))
            End If
        End If
        Return -1
    End Function

    Private Function ParseDiceChance(ByVal Text As String, ByVal DiceFaces As Integer) As Decimal
        If Text.Contains("+") And Text.Contains("-") Then
            Return -1
        ElseIf Text.Contains("+") Then
            Dim SplitText() As String = Text.Split("+")
            If SplitText.Count > 2 Or SplitText(1) <> "" Then
                Return -1
            End If
            If IsNumeric(SplitText(0)) AndAlso Decimal.Parse(SplitText(0)) <= DiceFaces + 1 AndAlso Decimal.Parse(SplitText(0)) >= 1 Then
                Return (DiceFaces - Decimal.Parse(SplitText(0)) + 1) / DiceFaces
            Else
                Return -1
            End If
        Else
            Dim SplitText() As String = Text.Split("-")
            If SplitText.Count > 2 Or SplitText(1) <> "" Then
                Return -1
            End If
            If IsNumeric(SplitText(0)) AndAlso Decimal.Parse(SplitText(0)) <= DiceFaces AndAlso Decimal.Parse(SplitText(0)) >= 0 Then
                Return Decimal.Parse(SplitText(0)) / DiceFaces
            Else
                Return -1
            End If
        End If
    End Function

    Private Sub txtNumDice_DoubleClick(sender As Object, e As EventArgs) Handles txtNumDice.DoubleClick
        'Default value for txtNumDice = "10"
        txtNumDice.Text = "10"
    End Sub

    Private Sub TextChangedNoFraction(sender As Object, e As EventArgs) Handles txtNumExtraHits.TextChanged
        'Checks for non-zero positive integer values on some textboxes. Rounds decimals up.
        Dim CurrentTextBox As TextBox = DirectCast(sender, TextBox)
        If IsNumeric(CurrentTextBox.Text) = False Then
            CurrentTextBox.BackColor = Color.Yellow
        ElseIf Format(CurrentTextBox.Text) <= 0 Then
            CurrentTextBox.BackColor = Color.Yellow
        Else
            CurrentTextBox.Text = Math.Ceiling(Double.Parse(CurrentTextBox.Text))
            CurrentTextBox.BackColor = Color.White
        End If
    End Sub

    Private Sub TextChangedFraction(sender As Object, e As EventArgs) Handles txtHitChance.TextChanged, txtWoundChance.TextChanged, txtSaveChance.TextChanged, txtFNPChance.TextChanged, txtAutoWoundChance.TextChanged, txtModifiedAPChance.TextChanged, txtModifiedSaveChance.TextChanged, txtMWWoundChance.TextChanged, txtExtraHits.TextChanged, txtHitRerollChance.TextChanged, txtWoundRerollChance.TextChanged, txtSaveRerollChance.TextChanged, txtFNPRerollChance.TextChanged
        'Checks for numeric values between 0 and 1 for probabilities, and automatically parses fracitons using ParseFraction
        Dim CurrentTextBox As TextBox = DirectCast(sender, TextBox)
        If (Not CurrentTextBox.Text.Contains("+")) AndAlso Not (CurrentTextBox.Text.Contains("-")) AndAlso IsNumeric(CurrentTextBox.Text) Then
            If Format(CurrentTextBox.Text) >= 0 And Format(CurrentTextBox.Text) <= 1 Then
                CurrentTextBox.BackColor = Color.White
                CurrentTextBox.Tag = ""
            Else
                CurrentTextBox.BackColor = Color.Yellow
            End If
        Else
            Dim ConvertedFraction As Decimal = ParseProbability(CurrentTextBox.Text, GetDiceSize(cmbDiceType.SelectedIndex))
            If ConvertedFraction <> -1 Then
                CurrentTextBox.Tag = ConvertedFraction.ToString
                CurrentTextBox.BackColor = Color.White
            Else
                CurrentTextBox.BackColor = Color.Yellow
                CurrentTextBox.Tag = ""
            End If
        End If
    End Sub

    Private Sub TextDefaultFractionHalf(sender As Object, e As EventArgs) Handles txtHitChance.DoubleClick, txtWoundChance.DoubleClick, txtSaveChance.DoubleClick
        'Core probabilities are set to 0.5
        DirectCast(sender, TextBox).Text = "0.5"
    End Sub

    Private Sub TextDefaultFractionZero(sender As Object, e As EventArgs) Handles txtAutoWoundChance.DoubleClick, txtModifiedAPChance.DoubleClick, txtMWWoundChance.DoubleClick, txtModifiedSaveChance.DoubleClick, txtExtraHits.DoubleClick, txtHitRerollChance.DoubleClick, txtWoundRerollChance.DoubleClick, txtSaveRerollChance.DoubleClick, txtFNPRerollChance.DoubleClick, txtRandomDamage.DoubleClick, txtRandomAttacks.DoubleClick, txtRerollAttacksChance.DoubleClick, txtRerollDamageChance.DoubleClick
        'Optional probabilities are set to 0
        DirectCast(sender, TextBox).Text = "0"
    End Sub

    Private Sub TextFillFraction(sender As Object, e As EventArgs) Handles txtHitChance.LostFocus, txtWoundChance.LostFocus, txtSaveChance.LostFocus, txtFNPChance.LostFocus, txtAutoWoundChance.LostFocus, txtModifiedAPChance.LostFocus, txtModifiedSaveChance.LostFocus, txtMWWoundChance.LostFocus, txtExtraHits.LostFocus, txtHitRerollChance.LostFocus, txtWoundRerollChance.LostFocus, txtSaveRerollChance.LostFocus, txtFNPRerollChance.LostFocus
        Dim CurrentTextBox As TextBox = DirectCast(sender, TextBox)
        If CurrentTextBox.Tag <> "" And IsNumeric(CurrentTextBox.Tag) Then
            CurrentTextBox.Text = Math.Round(Decimal.Parse(CurrentTextBox.Tag), 12).ToString
            CurrentTextBox.Tag = ""
        ElseIf IsNumeric(CurrentTextBox.Text) Then
            CurrentTextBox.Text = Math.Round(Decimal.Parse(CurrentTextBox.Text), 12).ToString
            CurrentTextBox.Tag = ""
        Else
            CurrentTextBox.Tag = ""
            CurrentTextBox.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

        'Check textboxes for valid entry
        Dim CheckEntries As Boolean = True
        Dim ErrorMessage As String = ""
        For Each Entry As Object In Me.Controls
            If TypeOf (Entry) Is TextBox Then
                If Entry.Name <> "txtNumDice" And Entry.Name <> "txtDamage" And Entry.Name <> "txtNumExtraHits" And Entry.Name <> "txtRandomDamage" And Entry.name <> "txtWoundNumMW" And Entry.Name <> "txtHitNumMW" And Entry.Name <> "txtRandomAttacks" Then
                    'Most textboxes require probabilities (numbers between 0 and 1 inclusive)
                    If IsNumeric(Entry.text) = False Then
                        CheckEntries = False
                    ElseIf Format(Entry.text) < 0 Then
                        CheckEntries = False
                    ElseIf Format(Entry.text) > 1 Then
                        CheckEntries = False
                    End If
                ElseIf Entry.name = "txtNumExtraHits" Or Entry.Name = "txtHitNumMW" Or Entry.Name = "txtWoundNumMW" Then
                    'txtNumExtraHits, txtHitNumMW and txtWoundNumMW require numerical entries greater than 0
                    If IsNumeric(Entry.text) = False Then
                        CheckEntries = False
                    ElseIf Format(Entry.text) <= 0 Then
                        CheckEntries = False
                    End If
                ElseIf Entry.name = "txtRandomDamage" Or Entry.Name = "txtNumDice" Or Entry.Name = "txtRandomAttacks" Then
                    'txtRandomDamage, txtNumDice and txtRandomAttacks require numerical entries equal to or greater than 0
                    If IsNumeric(Entry.text) = False Then
                        CheckEntries = False
                    ElseIf Format(Entry.text) < 0 Then
                        CheckEntries = False
                    End If
                ElseIf Entry.Name = "txtDamage" Then
                    'txtDamage require numerical integer entries
                    If IsNumeric(Entry.text) = False Then
                        CheckEntries = False
                    End If
                ElseIf Trim(Entry.Text) <> "" Then
                    If IsNumeric(Entry.text) = False Then
                        CheckEntries = False
                    ElseIf Format(Entry.text) <= 0 Then
                        CheckEntries = False
                    End If
                End If
            End If
        Next

        If CheckEntries = False Then
            ErrorMessage = "Please correct invalid entries (higlighted yellow). "
        End If

        If Format(txtDamage.Text) = 0 And Format(txtRandomDamage.Text) = 0 Then
            CheckEntries = False
            ErrorMessage += "At least one of Fixed/Random Damage must be non-zero. "
        End If

        If Format(txtNumDice.Text) = 0 And Format(txtRandomAttacks.Text) = 0 Then
            CheckEntries = False
            ErrorMessage += "At least one of Number/Random Attacks must be non-zero. "
        End If

        If Format(txtAutoWoundChance.Text) <> 0 And Format(txtExtraHits.Text) <> 0 And Format(txtAutoWoundChance.Text) <> Format(txtExtraHits.Text) Then
            CheckEntries = False
            ErrorMessage += "Lethal Hits and Sustained Hits must be the same if non-zero. "
        End If

        If Double.Parse(txtHitChance.Text) < Math.Max(Double.Parse(txtAutoWoundChance.Text), Double.Parse(txtExtraHits.Text)) Then
            CheckEntries = False
            txtHitChance.Text = Format(Math.Max(Double.Parse(txtAutoWoundChance.Text), Double.Parse(txtExtraHits.Text)))
            ErrorMessage += "Hit Chance cannot be lower than Crit chance. The values have been corrected. "
        End If

        If Double.Parse(txtWoundChance.Text) < Math.Max(Double.Parse(txtMWWoundChance.Text), Double.Parse(txtModifiedAPChance.Text)) Then
            CheckEntries = False
            txtWoundChance.Text = Format(Math.Max(Double.Parse(txtMWWoundChance.Text), Double.Parse(txtModifiedAPChance.Text)))
            ErrorMessage += "Wound Chance cannot be lower than Crit chance. The values have been corrected. "
        End If

        If Double.Parse(txtModifiedAPChance.Text) > 0 And Double.Parse(txtMWWoundChance.Text) > 0 Then
            CheckEntries = False
            txtModifiedAPChance.Text = "0"
            ErrorMessage += "Cannot have both modified AP and Devastating wounds. Modified AP chance has been removed. "
        End If

        If Integer.Parse(txtRandomAttacks.Text) > 20 Or Integer.Parse(txtRandomDamage.Text) > 20 Then
            CheckEntries = False
            ErrorMessage += "Random attacks and damage cannot have more than 20 dice. "
        End If

        'Display Message Box with acculumated error messages, then exit sub
        If CheckEntries = False Then
            MsgBox(ErrorMessage, vbOKOnly, "Data Error")
            Exit Sub
        End If

        '****Number of Attacks****

        Dim CurrentDist() As Double
        CurrentDist = DiceSumPDF(Integer.Parse(txtNumDice.Text), Integer.Parse(txtRandomAttacks.Text), 3 * cmbRandomAttackType.SelectedIndex + 3)

        '****Hit Chance****

        'Calculate chances corrected for rerolls. It is assumed that the lowest dice will be rerolled, even if this is successful
        'This could be fishing for critical hits, for example
        Dim RerollHitChance As Double = CalculateHitChance(Double.Parse(txtHitChance.Text), Math.Max(Double.Parse(txtAutoWoundChance.Text), Double.Parse(txtExtraHits.Text)), Double.Parse(txtHitRerollChance.Text))
        Dim RerollCritChance As Double = CalculateCritChance(Math.Max(Double.Parse(txtAutoWoundChance.Text), Double.Parse(txtExtraHits.Text)), Double.Parse(txtHitRerollChance.Text))


        'Maximum possible successful hits including all sustained hits, but subtracting lethal hits from the count of sustained hits.
        Dim MaxHits As Integer = (CurrentDist.Length - 1) * (1 + (Format(txtNumExtraHits.Text) - Math.Ceiling(Decimal.Parse(txtAutoWoundChance.Text))) * Math.Ceiling(Decimal.Parse(txtExtraHits.Text)))

        'Maximum possible wounds including wounds caused by lethal hits
        Dim MaxWounds As Integer = (CurrentDist.Length - 1) * (1 + Integer.Parse(txtNumExtraHits.Text) * Math.Ceiling(Decimal.Parse(txtExtraHits.Text)))

        'Temporary 3 dimensional array to track lethal hits distributions for every possible number of regular hits
        '1st dimension is distribution, 2nd is corresponding number of regular hits, 3rd is number of dice
        Dim LethalHitsPDFArray(CurrentDist.Length - 1, MaxHits, CurrentDist.Length - 1) As Double

        '2 dimensional array to track lethal hits distributions for every possible number of regular hits
        '1st dimension is distribution, 2nd is corresponding number of regular hits. Pr(Lethal | Regular)q
        Dim LethalHitsPDF(CurrentDist.Length - 1, MaxHits) As Double
        Dim MaxLethalHits As Integer = CurrentDist.Length - 1

        '1st dimension indicates number of resultant hits, 2nd dimension indicates starting number of dice
        Dim HitSumArray(MaxHits, CurrentDist.Length - 1) As Double



        If Format(txtAutoWoundChance.Text) > 0 And Format(txtExtraHits.Text) > 0 Then

            'Saving as variable for easy reference. This is only the extra hits as the original hit will trigger a lethal hit
            Dim SustainedNum As Integer = Integer.Parse(txtNumExtraHits.Text)

            'Represents the chance of having enough dice to achieve a particular result
            Dim TotalDiceChance(MaxHits) As Double

            'The calculation here is based on conditional probability:
            'Pr(Crits AND Hits) = Pr(Hits | Crits) * Pr(Crits)
            'Pr(x,y) = Binomial(TotalDice, Number Crits, Crit Probability) * Binomial(TotalDice - Number Crits, Number Hits, HitOnlyProbability/(1-Crit Probability))
            'These have to be added together for every combination of hits and crits that generates the same number of overall hits
            For DiceNum = 0 To CurrentDist.Length - 1

                'Skip calculations for combinations of dice that are impossible
                If CurrentDist(DiceNum) = 0 Then Continue For

                'Recalculate MaxHits for each potential number of dice. This will save on redundant calculations
                Dim CurrentMaxHits As Integer = DiceNum * SustainedNum
                Dim HitRow(CurrentMaxHits) As Double
                Dim CritDist() As Double = BinomPDF(DiceNum, RerollCritChance, "Sustained + Lethal Hits Calculation")
                For CritNum = 0 To DiceNum
                    'Pr(Hit|Crit)
                    Dim HitOnlyDist() As Double = BinomPDF(DiceNum - CritNum, RerollHitChance / (1 - RerollCritChance), "Sustained + Lethal Hits Calculation")
                    For HitNum = 0 To DiceNum - CritNum

                        'Probabilities are summed to the appropriate position in the array.
                        'This will be somewhat random, but will ensure that all relevant values for a BinomPDF
                        'without recalculating
                        HitRow(CritNum * SustainedNum + HitNum) += CritDist(CritNum) * HitOnlyDist(HitNum)
                        LethalHitsPDFArray(CritNum, CritNum * SustainedNum + HitNum, DiceNum) += CritDist(CritNum) * HitOnlyDist(HitNum)
                    Next
                Next

                'Add the distribution for the given number of dice to the overall distribution
                For TotalHitNum = 0 To HitRow.Length - 1
                    HitSumArray(TotalHitNum, DiceNum) = HitRow(TotalHitNum)
                    For CritNum = 0 To TotalHitNum / SustainedNum
                        If LethalHitsPDFArray(CritNum, TotalHitNum, DiceNum) <> 0 Then
                            LethalHitsPDFArray(CritNum, TotalHitNum, DiceNum) = LethalHitsPDFArray(CritNum, TotalHitNum, DiceNum) / HitRow(TotalHitNum)
                        End If
                    Next
                    If HitRow(TotalHitNum) = 0 Then Continue For
                    TotalDiceChance(TotalHitNum) += CurrentDist(DiceNum)
                Next
            Next

            'Totals up the Lethal Hits for each possible number of attack dice. Weighted by how likely that number of dice is to
            'appear out of the total numbers of dice capable of producing that result
            For TotalHitNum = 0 To MaxHits
                For DiceNum = 0 To CurrentDist.Length - 1
                    For CritNum = 0 To TotalHitNum / SustainedNum
                        If LethalHitsPDFArray(CritNum, TotalHitNum, DiceNum) <> 0 Then
                            LethalHitsPDF(CritNum, TotalHitNum) += CurrentDist(DiceNum) / TotalDiceChance(TotalHitNum) * LethalHitsPDFArray(CritNum, TotalHitNum, DiceNum)
                        End If
                    Next
                Next
            Next


        ElseIf Format(txtAutoWoundChance.Text) > 0 Then

            For DiceNum = 0 To CurrentDist.Length - 1

                'Skip calculations for combinations of dice that are impossible
                If CurrentDist(DiceNum) = 0 Then Continue For

                Dim HitDist() As Double = BinomPDF(DiceNum, RerollHitChance, "Hit Roll")
                For index = 0 To HitDist.Length - 1
                    HitSumArray(index, DiceNum) = HitDist(index)
                Next
            Next

            'Lethal hits calculated as conditional probability distribution based on regular hits
            For HitNum = 0 To MaxHits
                For DiceNum = HitNum To CurrentDist.Length - 1

                    'Skip calculations for combinations of dice that are impossible
                    'If CurrentDist(DiceNum) = 0 Then Continue For

                    Dim CritPDF() As Double = BinomPDF(DiceNum - HitNum, RerollCritChance / (1 - RerollHitChance), "Lethal Hits Calculation")
                    For CritNum = 0 To DiceNum - HitNum
                        LethalHitsPDF(CritNum, HitNum) += CurrentDist(DiceNum) * CritPDF(CritNum)
                    Next
                Next
            Next

        ElseIf Format(txtExtraHits.Text) > 0 Then

            'Saving as variable for easy reference. +1 is required because crits generate on top of the existing hit.
            'This is different if the crit also trigers lethal hits
            Dim SustainedNum As Integer = Integer.Parse(txtNumExtraHits.Text) + 1

            'The calculation here is based on conditional probability:
            'Pr(x,y) = Binomial(TotalDice, Number Crits, Crit Probability) * Binomial(TotalDice - Number Crits, Number Hits, HitOnlyProbability/(1-Crit Probability))
            'These have to be added together for every combination of hits and crits that generates the same number of overall hits
            For DiceNum = 0 To CurrentDist.Length - 1

                'Skip calculations for combinations of dice that are impossible
                If CurrentDist(DiceNum) = 0 Then Continue For

                'Recalculate MaxHits for each potential number of dice. This will save on redundant calculations
                Dim CurrentMaxHits As Integer = DiceNum * SustainedNum
                Dim HitRow(CurrentMaxHits) As Double
                Dim CritDist() As Double = BinomPDF(DiceNum, RerollCritChance, "Sustained Hits Only Calculation")
                For CritNum = 0 To DiceNum
                    Dim HitDist() As Double = BinomPDF(DiceNum - CritNum, RerollHitChance / (1 - RerollCritChance), "Sustained Hits Only Calculation")
                    For HitNum = 0 To CurrentMaxHits - SustainedNum * CritNum

                        'Must avoid solutions that require more dice than currently available
                        If CritNum + HitNum > DiceNum Then Continue For

                        'Probabilities are summed to the appropriate position in the array.
                        'This will be somewhat random, but will ensure that all relevant values for a BinomPDF
                        'without recalculating
                        HitRow(CritNum * SustainedNum + HitNum) += CritDist(CritNum) * HitDist(HitNum)
                    Next
                Next

                'Add the distribution for the given number of dice to the overall distribution
                For HitNum = 0 To HitRow.Length - 1
                    HitSumArray(HitNum, DiceNum) = HitRow(HitNum)
                Next
            Next

            'Lethal Hits PDF must show 100% probability for 0 lethal hits if not used
            For HitNum = 0 To MaxHits
                LethalHitsPDF(0, HitNum) = 1
            Next

        Else
            For DiceNum = 0 To CurrentDist.Length - 1

                'Skip calculations for combinations of dice that are impossible
                If CurrentDist(DiceNum) = 0 Then Continue For

                Dim HitDist() As Double = BinomPDF(DiceNum, RerollHitChance, "Hit Roll")
                For index = 0 To HitDist.Length - 1
                    HitSumArray(index, DiceNum) = HitDist(index)
                Next
            Next

            'Lethal Hits PDF must show 100% probability for 0 lethal hits if not used
            For HitNum = 0 To MaxHits
                LethalHitsPDF(0, HitNum) = 1
            Next
        End If

        'This applies the second distribution over the first. In this case the distribution of hits over the distribution of attack number
        Dim TempDist(MaxHits) As Double
        For AttackNum = 0 To MaxHits
            For index = 0 To CurrentDist.Length - 1
                TempDist(AttackNum) += CurrentDist(index) * HitSumArray(AttackNum, index)
            Next
        Next
        Array.Resize(CurrentDist, MaxHits + 1)
        CurrentDist = TempDist.Clone

        '****Wound Roll****

        'Recalculate hit/crit values for wound roll stage
        RerollCritChance = CalculateCritChance(Math.Max(Double.Parse(txtMWWoundChance.Text), Double.Parse(txtModifiedAPChance.Text)), Double.Parse(txtWoundRerollChance.Text))
        RerollHitChance = CalculateHitChance(Double.Parse(txtWoundChance.Text), Math.Max(Double.Parse(txtMWWoundChance.Text), Double.Parse(txtModifiedAPChance.Text)), Double.Parse(txtWoundRerollChance.Text))

        'Maximum possible wounds are calculated at hits step to account for lethal/sustained hits

        '1st dimension is wounds PDF, 2nd dimension is corresponding number of hits
        Dim WoundSumArray(MaxWounds, CurrentDist.Length - 1) As Double

        '1st dimension is AP PDF, 2nd Dimension is corresponding number of normal wounds, 3rd dimension is corresponding number of hits
        Dim ModifiedAPPDFArray(CurrentDist.Length - 1, MaxWounds, CurrentDist.Length - 1) As Double

        '1st dimension is AP PDF, 2nd dimension is corresponding number of normal wounds
        Dim ModifiedAPPDF(CurrentDist.Length - 1, MaxWounds) As Double

        'Save Maximum Crit Wounds for use in Save step
        Dim MaxModified As Integer = CurrentDist.Length - 1

        'Storing modified save chance. If using devastating wounds, then this will be set to 1 (instant fail)
        Dim ModifiedSaveChance As Double = Double.Parse(txtModifiedSaveChance.Text) - Double.Parse(txtSaveRerollChance.Text) * (1 - Double.Parse(txtModifiedSaveChance.Text))
        If Double.Parse(txtMWWoundChance.Text) > 0 Then ModifiedSaveChance = 1




        If Double.Parse(txtMWWoundChance.Text) > 0 Or Double.Parse(txtModifiedAPChance.Text) > 0 Then

            For HitNum = 0 To CurrentDist.Length - 1
                Dim WoundRow() As Double = BinomPDF(HitNum, RerollHitChance, "Wound Roll with Crits")
                For WoundNum = 0 To HitNum
                    Dim CritRow() As Double = BinomPDF(HitNum - WoundNum, RerollCritChance / (1 - RerollHitChance), "Wound Roll with Crits")
                    For AutoWoundNum = 0 To Math.Min(MaxLethalHits, MaxWounds - WoundNum)
                        'Calculate probability of Wounds And Lethal Hits for a given number of hits
                        WoundSumArray(AutoWoundNum + WoundNum, HitNum) += WoundRow(WoundNum) * LethalHitsPDF(AutoWoundNum, HitNum)
                        For CritNum = 0 To CritRow.Length - 1
                            'Calculate the probability of critical wounds AND a given number of total wounds and lethal hits.
                            'This is an intersection probability, not conditional (all possible combinations should sum to 1)
                            'Sum together the different possibilities of lethal hits for a particular combination of crits and wounds
                            ModifiedAPPDFArray(CritNum, AutoWoundNum + WoundNum, HitNum) += WoundRow(WoundNum) * CritRow(CritNum) * LethalHitsPDF(AutoWoundNum, HitNum)
                        Next
                    Next
                Next
            Next

            'Need to know the overall wounds PDF for the next step, so it is calculated early
            ReDim TempDist(MaxWounds)
            For WoundCount = 0 To MaxWounds
                For HitCount = 0 To CurrentDist.Length - 1
                    TempDist(WoundCount) += CurrentDist(HitCount) * WoundSumArray(WoundCount, HitCount)
                Next
            Next

            For WoundNum = 0 To MaxWounds
                For HitNum = 0 To MaxHits
                    For CritNum = 0 To MaxHits
                        'Sum together the proportional numbers of crits and wounds for a given number of hits
                        'This is divided by the probability of that particular number of wounds to find Pr (Crits | Wounds)
                        If TempDist(WoundNum) > 0 Then
                            ModifiedAPPDF(CritNum, WoundNum) += CurrentDist(HitNum) * ModifiedAPPDFArray(CritNum, WoundNum, HitNum) / TempDist(WoundNum)
                        End If
                    Next
                Next
            Next

        Else

            For HitNum = 0 To CurrentDist.Length - 1
                Dim WoundRow() As Double = BinomPDF(HitNum, RerollHitChance, "Wound Roll Step")

                'Just do a plain calculation if there are no lethal hits to include
                If Double.Parse(txtAutoWoundChance.Text) = 0 Then
                    For i = 0 To HitNum
                        WoundSumArray(i, HitNum) = WoundRow(i)
                    Next
                    Continue For
                End If
                'Copy out corresponding row of Lethal Hits Array
                Dim LethalHitRow(MaxLethalHits) As Double
                For i = 0 To MaxLethalHits
                    LethalHitRow(i) = LethalHitsPDF(i, HitNum)
                Next
                'Collect the PDF of a given number of wounds in total using the SumDistributions function
                Dim TotalWoundRow() As Double = SumDistributions(WoundRow, LethalHitRow, MaxWounds)
                For i = 0 To TotalWoundRow.Length - 1
                    WoundSumArray(i, HitNum) = TotalWoundRow(i)
                Next
            Next

            ReDim TempDist(MaxWounds)
            For WoundCount = 0 To MaxWounds
                For HitCount = 0 To CurrentDist.Length - 1
                    TempDist(WoundCount) += CurrentDist(HitCount) * WoundSumArray(WoundCount, HitCount)
                Next
            Next

        End If

        'Finally update CurrentDist to represent regular wounds (not modified AP or devastating wounds)
        Array.Resize(CurrentDist, TempDist.Length)
        CurrentDist = TempDist.Clone

        '****Save Step****

        'Save rerolls are different because a failed save is rerolled (which lowers the probability of a fail instead of increasing)
        Dim RerollSaveChance As Double = Double.Parse(txtSaveChance.Text) - Double.Parse(txtSaveRerollChance.Text) * (1 - Double.Parse(txtSaveChance.Text))
        'Modified save chance was calculated in the Wound step, and set to 1 if using devastating wounds

        '1st dimension is PDF, 2nd dimension is corresponding number of wounds
        Dim SavePDF(CurrentDist.Length - 1, CurrentDist.Length - 1) As Double

        'Same distribution but using modified ap save. Not used if using Devastating Wounds
        Dim ModifiedSavePDF(CurrentDist.Length - 1) As Double

        'SaveNum represents the number of *failed* saves
        For WoundNum = 0 To CurrentDist.Length - 1
            Dim SaveRow() As Double = BinomPDF(WoundNum, RerollSaveChance, "Normal Save Rolls Step")
            For SaveNum = 0 To WoundNum
                'Store the probability of getting a certain number of failed saves GIVEN a certain number of initial wounds
                SavePDF(SaveNum, WoundNum) += SaveRow(SaveNum)
            Next
        Next

        ReDim TempDist(CurrentDist.Length - 1)

        'Mortal Wound PDF has to be adjusted to reflect unsaved wounds.
        'This checks if using modified AP, or if using non-rollover Dev Wounds. 
        If (Decimal.Parse(txtModifiedAPChance.Text) > 0 And Decimal.Parse(txtMWWoundChance.Text) = 0) Or (Decimal.Parse(txtModifiedAPChance.Text) = 0 And Decimal.Parse(txtMWWoundChance.Text) > 0 And chkRolloverDevWounds.Checked = False) Then
            'For a given number of initial normal wounds

            'Set modified save chance to always fail if using non-rollover Dev Wounds
            If Decimal.Parse(txtModifiedAPChance.Text) = 0 And chkRolloverDevWounds.Checked = False Then ModifiedSaveChance = 1

            For WoundNum = 0 To CurrentDist.Length - 1
                'For a given number of wounds with modified AP
                For ModifiedWoundNum = 0 To MaxModified
                    'Skip any situations where it would be impossible to have that many modified wounds
                    If ModifiedAPPDF(ModifiedWoundNum, WoundNum) > 0 Then
                        Dim ModifiedRow() As Double = BinomPDF(ModifiedWoundNum, ModifiedSaveChance, "Modified AP Save Rolls Step")
                        'For a given number of normal failed saves
                        For SaveNum = 0 To WoundNum
                            'For a given number of failed modified saves
                            For ModifiedSaveNum = 0 To ModifiedWoundNum
                                'Pr(Failed Save | Normal Wounds) * Pr(Normal Wounds) * Pr(Modified Wounds | Normal Wounds) * Pr(Failed Modified Saves | Modified Wounds)
                                TempDist(SaveNum + ModifiedSaveNum) += SavePDF(SaveNum, WoundNum) * CurrentDist(WoundNum) * ModifiedAPPDF(ModifiedWoundNum, WoundNum) * ModifiedRow(ModifiedSaveNum)
                            Next
                        Next
                    End If
                Next
            Next
        Else

            Dim TempMWPDF(MaxModified, CurrentDist.Length - 1) As Double
            Dim SaveFactor(CurrentDist.Length - 1) As Double
            For WoundNum = 0 To CurrentDist.Length - 1
                For SaveNum = 0 To WoundNum
                    'Pr(Failed Save | Normal Wounds) * Pr(Normal Wounds)
                    TempDist(SaveNum) += SavePDF(SaveNum, WoundNum) * CurrentDist(WoundNum)
                Next
                For SaveNum = 0 To WoundNum
                    For MWNum = 0 To MaxModified
                        If TempDist(WoundNum) > 0 Then
                            TempMWPDF(MWNum, SaveNum) += CurrentDist(WoundNum) * SavePDF(SaveNum, WoundNum) * ModifiedAPPDF(MWNum, WoundNum)
                            SaveFactor(SaveNum) += CurrentDist(WoundNum) * SavePDF(SaveNum, WoundNum) * ModifiedAPPDF(MWNum, WoundNum)
                        End If
                    Next
                Next
            Next
            For SaveNum = 0 To CurrentDist.Length - 1
                For MwNum = 0 To MaxModified
                    If TempMWPDF(MwNum, SaveNum) > 0 Then
                        TempMWPDF(MwNum, SaveNum) = TempMWPDF(MwNum, SaveNum) / SaveFactor(SaveNum)
                    End If
                Next
            Next
            If Double.Parse(txtMWWoundChance.Text) > 0 Then
                ModifiedAPPDF = TempMWPDF.Clone
            End If
        End If

        CurrentDist = TempDist.Clone

        '****Wounds needed to inflict Causalty****
        'First generate a PDF representing the chance of a certain number of wounds being capable of inflicting a casualty
        'This will account for FeelNoPain tests

        Dim FNPChance As Double = Double.Parse(txtFNPChance.Text) - Double.Parse(txtFNPRerollChance.Text) * (1 - Double.Parse(txtFNPChance.Text))

        Dim RndDmg As Integer = cmbRandomDamageType.SelectedIndex * 3 + 3
        Dim MaxDamage As Integer = Integer.Parse(txtDamage.Text) + Integer.Parse(txtRandomDamage.Text) * RndDmg
        Dim FNPPDF(MaxDamage) As Double
        Dim RandomDamagePDF() As Double

        ''Generate PDF of probabilities of different amounts of damage from a single unsaved wound
        ''Includes FNP saves

        RandomDamagePDF = DiceSumPDF(Integer.Parse(txtDamage.Text), Integer.Parse(txtRandomDamage.Text), 3 * cmbRandomDamageType.SelectedIndex + 3)

        If FNPChance < 1 Then
            For Damage = 1 To RandomDamagePDF.Length - 1
                Dim BinomRow() As Double = BinomPDF(Damage, FNPChance, "Casualties FNP Calculation")
                For UnsavedDamage = 0 To Damage
                    FNPPDF(UnsavedDamage) += RandomDamagePDF(Damage) * BinomRow(UnsavedDamage)
                Next
            Next
        Else
            'Do not alter the PDF if there is no FNP save
            FNPPDF = RandomDamagePDF.Clone
        End If

        If cmbTargetWounds.SelectedIndex > 0 Then

            rdbCasualties.Enabled = True
            rdbCasualties.Select()

            Wounds = cmbTargetWounds.SelectedIndex

            Dim MaxCasualties As Integer = Math.Floor(MaxDamage / Wounds)

            'Iterate through possible numbers of wound rolls and calculate the probabilities of each
            'FNP save gives the possibility of an infinite number of rolls, so this is capped at the maximum number of unsaved wounds
            '1st dimension is the casualty number, 2nd is the health
            Dim CasualtyWounds As List(Of Double()) = New List(Of Double())

            'It will always be impossible for a model to die with 0 unsaved wounds, so the first index is filled with a 100% probability of 0 Casualties 0 Wounds
            Dim CasualtyRowZero(MaxDamage * MaxWounds) As Double
            Dim UpperCasualties As Integer = 0
            CasualtyRowZero(0) = 1
            CasualtyWounds.Add(CasualtyRowZero)

            'While there is a possibility of needing additional wound rolls, and there are additional possible wounds left to apply damage
            For WoundCount = 1 To CurrentDist.Length - 1
                Dim CasualtyNum As Integer = 0
                Dim CasualtyRow(Math.Ceiling(MaxDamage * MaxWounds / Wounds) * Wounds) As Double
                While CasualtyNum <= UpperCasualties

                    For Damage = 0 To MaxDamage
                        'Apply damage distribution multiplied by the probability that the previous model was just killed
                        'The minimum is to add all probabilities to the damage that corresponds to a killed model, regardless of any excess damage
                        CasualtyRow(Math.Min(Damage, Wounds) + CasualtyNum * Wounds) += FNPPDF(Damage) * CasualtyWounds(WoundCount - 1)(CasualtyNum * Wounds)
                        'Apply the additional damage to the model currently being damaged. 
                        For Health = 1 To Wounds - 1
                            CasualtyRow(Math.Min(Health + Damage, Wounds) + CasualtyNum * Wounds) += FNPPDF(Damage) * CasualtyWounds(WoundCount - 1)(CasualtyNum * Wounds + Health)
                        Next
                    Next
                    'Increment to damaging the next model
                    CasualtyNum += 1
                End While
                'Check if it is possible to damage the next model on the next wound roll. If so increment the maximum number of casualties to count to
                If CasualtyRow(CasualtyNum * Wounds) > 0 Then
                    UpperCasualties += 1
                End If
                CasualtyWounds.Add(CasualtyRow)
            Next

            'Temporarily store the number of casualties without applying mortal wounds (Devastating Wounds)
            Dim NoMWPDF(Math.Ceiling(MaxDamage * MaxWounds / Wounds) * Wounds, CurrentDist.Length - 1) As Double

            For WoundNum = 0 To CurrentDist.Length - 1
                For DamageNum = 0 To CasualtyWounds(WoundNum).Length - 1
                    NoMWPDF(DamageNum, WoundNum) += CurrentDist(WoundNum) * CasualtyWounds(WoundNum)(DamageNum)
                Next
            Next


            'Apply mortal wounds if necessary. Skip this if not using rollover Dev Wounds
            If Double.Parse(txtMWWoundChance.Text) > 0 And chkRolloverDevWounds.Checked = True Then

                'Condense all possible numbers of mortal wounds into a single PDF array Pr(MW | W)
                Dim MWPDFArray(MaxModified * MaxDamage, CurrentDist.Length - 1) As Double
                Dim RunningDamagePDF() As Double = New Double() {1}
                For WoundNum = 0 To CurrentDist.Length - 1
                    For MWNum = 0 To MaxModified
                        For DamageNum = 0 To RunningDamagePDF.Length - 1
                            MWPDFArray(DamageNum, WoundNum) += RunningDamagePDF(DamageNum) * ModifiedAPPDF(MWNum, WoundNum)
                        Next
                        RunningDamagePDF = SumDistributions(RunningDamagePDF, FNPPDF, FNPPDF.Length + RunningDamagePDF.Length - 2)
                    Next
                    RunningDamagePDF = New Double() {1}
                Next

                ReDim CasualtiesPDF(MaxDamage * (MaxWounds + MaxModified))
                'For each unsaved wound
                For Woundnum = 0 To CurrentDist.Length - 1
                    'For each possible damage inflicted from all mortal wounds
                    For DamageNum = 0 To MaxDamage * MaxModified
                        'For each possible number of casualties already accrued
                        For HealthNum = 0 To MaxDamage * MaxWounds
                            CasualtiesPDF(HealthNum + DamageNum) += NoMWPDF(HealthNum, Woundnum) * MWPDFArray(DamageNum, Woundnum)
                        Next
                    Next
                Next
            Else
                ReDim CasualtiesPDF(MaxDamage * MaxWounds)
                For Woundnum = 0 To CurrentDist.Length - 1
                    For Healthnum = 0 To MaxDamage * MaxWounds
                        CasualtiesPDF(Healthnum) += NoMWPDF(Healthnum, Woundnum)
                    Next
                Next
            End If
        Else
            'If using single target, do not allow "casualties" to be graphed
            rdbCasualties.Enabled = False
            rdbWoundsLost.Select()

        End If

        Dim TempDamagePDF(Math.Max(MaxWounds, MaxModified) * MaxDamage, Math.Max(MaxWounds, MaxModified)) As Double

        'Calculate damage distribution (not accounting for model wounds characteristic)
        For WoundNum = 0 To Math.Max(MaxWounds, MaxModified)
            Dim PrevDamagePDF(0) As Double
            '100% chance of 0 damage from 0 unsaved wounds
            PrevDamagePDF(0) = 1
            'Iterate through the damage rolls
            For RollNum = 1 To WoundNum
                Dim NextDamagePDF(RollNum * MaxDamage) As Double
                For HealthNum = 0 To (RollNum - 1) * MaxDamage
                    For DamageNum = 0 To MaxDamage
                        'Add on the damage of the current roll to the previously accrued damage
                        If DamageNum + HealthNum <= RollNum * MaxDamage Then
                            NextDamagePDF(HealthNum + DamageNum) += FNPPDF(DamageNum) * PrevDamagePDF(HealthNum)
                        End If
                    Next
                Next
                PrevDamagePDF = NextDamagePDF.Clone()
            Next
            For index = 0 To PrevDamagePDF.Length - 1
                TempDamagePDF(index, WoundNum) = PrevDamagePDF(index)
            Next
        Next

        'Add on Mortal Wounds from Devastating Wounds. Skip this if not using rollover Dev Wounds
        If Double.Parse(txtMWWoundChance.Text) > 0 And chkRolloverDevWounds.Checked = True Then
            ReDim DamagePDF((MaxWounds + MaxModified) * MaxDamage)
            'For each unsaved wound
            For WoundNum = 0 To CurrentDist.Length - 1
                'For each Mortal Wound
                For MWNum = 0 To MaxModified
                    'For each possible damage value of an unsaved wound
                    For WoundDamageNum = 0 To WoundNum * MaxDamage
                        'For each possible damagevalue of a Mortal Wound
                        For MWDamageNum = 0 To MWNum * MaxDamage
                            DamagePDF(MWDamageNum + WoundDamageNum) += CurrentDist(WoundNum) * TempDamagePDF(WoundDamageNum, WoundNum) * ModifiedAPPDF(MWNum, WoundNum) * TempDamagePDF(MWDamageNum, MWNum)
                        Next
                    Next
                Next
            Next
        Else
            ReDim DamagePDF(MaxWounds * MaxDamage)
            For WoundNum = 0 To CurrentDist.Length - 1
                For DamageNum = 0 To MaxDamage * WoundNum
                    DamagePDF(DamageNum) += CurrentDist(WoundNum) * TempDamagePDF(DamageNum, WoundNum)
                Next
            Next
        End If

        '****Fill Spreadsheet****
        dgvDistributionSpreadsheet.Rows.Clear()
        Dim CumulativeProbability As Double = 1
        Dim ExpectedValue As Double = 0
        Dim StandardDeviation As Double
        If rdbCasualties.Checked = True Then
            Dim LabelList(CasualtiesPDF.Length - 1) As String
            dgvDistributionSpreadsheet.Columns(0).HeaderText = "Casualties - Damage"
            dgvDistributionSpreadsheet.Columns(0).Width = TextSize("Casualties - Damage", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True).Width + 10
            dgvDistributionSpreadsheet.ColumnHeadersHeight = 4
            For Each dgvColumn As DataGridViewColumn In dgvDistributionSpreadsheet.Columns
                If (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then
                    dgvDistributionSpreadsheet.ColumnHeadersHeight = (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY
                End If
            Next
            'Also creates a list of column labels for the histogram. Calculates the expected value
            For index = 0 To CasualtiesPDF.Length - 1
                dgvDistributionSpreadsheet.Rows.Add(New String() {Math.Floor(index / Wounds).ToString + " - " + (index Mod Wounds).ToString, Math.Round(CasualtiesPDF(index), 9).ToString, Math.Round(CumulativeProbability, 9).ToString})
                LabelList(index) = Math.Floor(index / Wounds).ToString + " - " + (index Mod Wounds).ToString
                CumulativeProbability -= CasualtiesPDF(index)
                ExpectedValue += index * CasualtiesPDF(index)
            Next
            'Expected value is required to calculate the standard deviation, so it is necessary to iterate through again
            For index = 0 To CasualtiesPDF.Length - 1
                StandardDeviation += (index - ExpectedValue) ^ 2 * CasualtiesPDF(index)
            Next
            'Correct Expected value to casualties - wounds and convert variance to standard deviation
            lblExpectedValue.Text = Math.Floor(ExpectedValue / Wounds).ToString + " - " + Math.Round((ExpectedValue Mod Wounds), 2).ToString
            lblStandardDeviation.Text = Math.Sqrt(StandardDeviation).ToString
            SelectedColumns.Clear()
            ColumnArray = DrawDistribution(picDistributionHistogram, CasualtiesPDF, LabelList)
        Else
            Dim LabelList(DamagePDF.Length - 1) As String
            dgvDistributionSpreadsheet.Columns(0).HeaderText = "Damage"
            dgvDistributionSpreadsheet.Columns(0).Width = TextSize("Damage", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True).Width + 10
            dgvDistributionSpreadsheet.ColumnHeadersHeight = 4
            For Each dgvColumn As DataGridViewColumn In dgvDistributionSpreadsheet.Columns
                If (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then
                    dgvDistributionSpreadsheet.ColumnHeadersHeight = (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY
                End If
            Next
            'Also creates a list of column labels for the histogram. Calculates the expected value
            For index = 0 To DamagePDF.Length - 1
                dgvDistributionSpreadsheet.Rows.Add(New String() {index.ToString, Math.Round(DamagePDF(index), 9).ToString, Math.Round(CumulativeProbability, 9).ToString})
                LabelList(index) = index.ToString
                ExpectedValue += index * DamagePDF(index)
                CumulativeProbability -= DamagePDF(index)
            Next
            'Expected value is required to calculate the standard deviation, so it is necessary to iterate through again
            For index = 0 To DamagePDF.Length - 1
                StandardDeviation += (index - ExpectedValue) ^ 2 * DamagePDF(index)
            Next
            'Convert variance to standard deviation and display both summary statistics
            lblExpectedValue.Text = ExpectedValue.ToString
            lblStandardDeviation.Text = Math.Sqrt(StandardDeviation).ToString
            SelectedColumns.Clear()
            ColumnArray = DrawDistribution(picDistributionHistogram, DamagePDF, LabelList)
        End If

    End Sub

    Private Sub TextDefaultOne(sender As Object, e As EventArgs) Handles txtNumExtraHits.DoubleClick, txtDamage.DoubleClick
        'Default for number of extra hits, Damage or MW is 1
        DirectCast(sender, TextBox).Text = "1"
    End Sub

    Private Sub txtDamage_TextChanged(sender As Object, e As EventArgs) Handles txtDamage.TextChanged
        'Checks for non-zero integer values on some textboxes. Rounds decimals up.
        'Negative values allow for damage reduction effects
        If IsNumeric(txtDamage.Text) = False Then
            txtDamage.BackColor = Color.Yellow
        Else
            txtDamage.Text = Math.Ceiling(Double.Parse(txtDamage.Text))
            txtDamage.BackColor = Color.White
        End If
    End Sub

    Private Sub TextChangedNoFractionZero(sender As Object, e As EventArgs) Handles txtRandomDamage.TextChanged, txtRandomAttacks.TextChanged, txtNumDice.TextChanged
        'Checks for positive or zero integer on some textboxes. Decimals are rounded up
        If IsNumeric(DirectCast(sender, TextBox).Text) = False Then
            DirectCast(sender, TextBox).BackColor = Color.Yellow
        ElseIf Format(DirectCast(sender, TextBox).Text) < 0 Then
            DirectCast(sender, TextBox).BackColor = Color.Yellow
        Else
            DirectCast(sender, TextBox).Text = Math.Ceiling(Double.Parse(DirectCast(sender, TextBox).Text))
            DirectCast(sender, TextBox).BackColor = Color.White
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using g As Graphics = Me.CreateGraphics()
            ScaleX = g.DpiX / 96
            ScaleY = g.DpiY / 96
        End Using

        dgvDistributionSpreadsheet.RowTemplate.Height = 20 * ScaleY

        cmbRandomDamageType.SelectedIndex = 0
        cmbRandomAttackType.SelectedIndex = 0
        cmbTargetWounds.SelectedIndex = 0
        Dim HeaderSize = TextSize("Damage", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True)
        Dim ResultColumn As New DataGridViewColumn
        ResultColumn.DataPropertyName = "Result"
        ResultColumn.HeaderText = "Damage"
        ResultColumn.Name = "colResult"
        ResultColumn.CellTemplate = New DataGridViewTextBoxCell
        ResultColumn.Width = HeaderSize.Width + 10
        If (HeaderSize.Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then dgvDistributionSpreadsheet.ColumnHeadersHeight = (HeaderSize.Height + 2) * ScaleY
        dgvDistributionSpreadsheet.Columns.Add(ResultColumn)

        HeaderSize = TextSize("Probability", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True)
        Dim ProbabilityColumn As New DataGridViewColumn
        ProbabilityColumn.DataPropertyName = "Probability"
        ProbabilityColumn.HeaderText = "Probability"
        ProbabilityColumn.Name = "colProbability"
        ProbabilityColumn.CellTemplate = New DataGridViewTextBoxCell
        ProbabilityColumn.Width = HeaderSize.Width + 10
        If (HeaderSize.Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then dgvDistributionSpreadsheet.ColumnHeadersHeight = (HeaderSize.Height + 2) * ScaleY
        dgvDistributionSpreadsheet.Columns.Add(ProbabilityColumn)

        HeaderSize = TextSize("Cumulative Probability", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True)
        Dim CumulativeColumn As New DataGridViewColumn
        CumulativeColumn.DataPropertyName = "Cumulative Probability"
        CumulativeColumn.HeaderText = "Cumulative Probability"
        CumulativeColumn.Name = "colCumulative"
        CumulativeColumn.CellTemplate = New DataGridViewTextBoxCell
        CumulativeColumn.Width = HeaderSize.Width + 10
        If (HeaderSize.Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then dgvDistributionSpreadsheet.ColumnHeadersHeight = (HeaderSize.Height + 2) * ScaleY
        dgvDistributionSpreadsheet.Columns.Add(CumulativeColumn)

        cmbDiceType.SelectedIndex = 1

    End Sub

    Private Sub txtFNPChance_DoubleClick(sender As Object, e As EventArgs) Handles txtFNPChance.DoubleClick
        'FNP Chance is set to 1 representing a 100% chance to FAIL a FNP test
        txtFNPChance.Text = "1"
    End Sub

    Private Sub ChangePDFDisplay(sender As Object, e As EventArgs) Handles rdbWoundsLost.CheckedChanged, rdbCasualties.CheckedChanged
        If DamagePDF Is Nothing OrElse DamagePDF.Length = 0 OrElse CasualtiesPDF Is Nothing OrElse CasualtiesPDF.Length = 0 Then Exit Sub
        dgvDistributionSpreadsheet.Rows.Clear()
        Dim CumulativeProbability As Double = 1
        Dim ExpectedValue As Double = 0
        Dim StandardDeviation As Double = 0
        If rdbCasualties.Checked = True Then
            Dim LabelList(CasualtiesPDF.Length - 1) As String
            dgvDistributionSpreadsheet.Columns(0).HeaderText = "Casualties - Damage"
            dgvDistributionSpreadsheet.Columns(0).Width = TextSize("Casualties - Damage", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True).Width + 10
            dgvDistributionSpreadsheet.ColumnHeadersHeight = 4
            For Each dgvColumn As DataGridViewColumn In dgvDistributionSpreadsheet.Columns
                If (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then
                    dgvDistributionSpreadsheet.ColumnHeadersHeight = (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY
                End If
            Next
            For index = 0 To CasualtiesPDF.Length - 1
                dgvDistributionSpreadsheet.Rows.Add(New String() {Math.Floor(index / Wounds).ToString + " - " + (index Mod Wounds).ToString, Math.Round(CasualtiesPDF(index), 9).ToString, Math.Round(CumulativeProbability, 9).ToString})
                LabelList(index) = Math.Floor(index / Wounds).ToString + " - " + (index Mod Wounds).ToString
                ExpectedValue += index * CasualtiesPDF(index)
                CumulativeProbability -= CasualtiesPDF(index)
            Next
            For index = 0 To CasualtiesPDF.Length - 1
                StandardDeviation += (index - ExpectedValue) ^ 2 * CasualtiesPDF(index)
            Next
            lblExpectedValue.Text = Math.Floor(ExpectedValue / Wounds).ToString + " - " + Math.Round(ExpectedValue Mod Wounds, 2).ToString
            lblStandardDeviation.Text = Math.Sqrt(StandardDeviation).ToString
            SelectedColumns.Clear()
            ColumnArray = DrawDistribution(picDistributionHistogram, CasualtiesPDF, LabelList)
        Else
            Dim LabelList(DamagePDF.Length - 1) As String
            dgvDistributionSpreadsheet.Columns(0).HeaderText = "Damage"
            dgvDistributionSpreadsheet.Columns(0).Width = TextSize("Damage", dgvDistributionSpreadsheet.DefaultCellStyle.Font, True).Width + 10
            dgvDistributionSpreadsheet.ColumnHeadersHeight = 4
            For Each dgvColumn As DataGridViewColumn In dgvDistributionSpreadsheet.Columns
                If (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY > dgvDistributionSpreadsheet.ColumnHeadersHeight Then
                    dgvDistributionSpreadsheet.ColumnHeadersHeight = (TextSize(dgvColumn.HeaderText, dgvColumn.DefaultCellStyle.Font, True).Height + 2) * ScaleY
                End If
            Next
            For index = 0 To DamagePDF.Length - 1
                dgvDistributionSpreadsheet.Rows.Add(New String() {index.ToString, Math.Round(DamagePDF(index), 9).ToString, Math.Round(CumulativeProbability, 9).ToString})
                LabelList(index) = index.ToString
                ExpectedValue += index * DamagePDF(index)
                CumulativeProbability -= DamagePDF(index)
            Next
            For index = 0 To DamagePDF.Length - 1
                StandardDeviation += (index - ExpectedValue) ^ 2 * DamagePDF(index)
            Next
            lblExpectedValue.Text = ExpectedValue.ToString
            lblStandardDeviation.Text = Math.Sqrt(StandardDeviation).ToString
            SelectedColumns.Clear()
            ColumnArray = DrawDistribution(picDistributionHistogram, DamagePDF, LabelList)
        End If
    End Sub

    'Returns an array of the rectangles of each column in the pdf graph
    'Distribution is the array of instantaneous probabilities for a given damage index
    'XArray is the array of labels for the damage indices
    Private Function DrawDistribution(Canvas As PictureBox, Distribution() As Double, XArray() As String) As List(Of Rectangle)

        If Distribution.Count <> XArray.Count Then Return Nothing



        'Store the minimum and maximum displayed members of the PDF.
        'These will be the smallest and largest values with a meaningful chance of occuring
        Dim DistMin As Integer = 0
        Dim DistMax As Integer = Distribution.Length - 1
        Dim YScale As Double = -(Canvas.Height - 100 * ScaleY) / Distribution.Max()
        Dim index As Integer = 0

        'Iterate through the distribution from either end, and record the first index with a probability >= 0.001
        While index < Distribution.Length - 1 AndAlso Distribution(index) < 0.001
            index += 1
            DistMin = index
        End While
        index = Distribution.Length - 1
        While index >= DistMin And Distribution(index) < 0.001
            index -= 1
            DistMax = index
        End While

        'Create a map of the x coordinates to the corresponding member of the PDF
        Dim GraphLength As Integer = DistMax - DistMin
        Dim xmap(Canvas.Width) As Integer
        For x = 0 To Canvas.Width
            xmap(x) = Math.Max((x / Canvas.Width) * (GraphLength + 1) - 0.5, 0) + DistMin
        Next

        'Map out the rectangles of each column in the displayed graph.
        'Any columns not displayed are set to 0,0,0,0
        Dim GraphArray As List(Of Rectangle) = New List(Of Rectangle)
        Dim ColumnNum As Integer = 0
        Dim CurrentX As Integer = 0
        For index = 0 To Distribution.Length - 1
            Dim NextColumn As Rectangle = New Rectangle(New Point(CurrentX, 0), New Drawing.Size(0, Canvas.Height - 100 * ScaleY))
            While CurrentX <= Canvas.Width AndAlso xmap(CurrentX) = index
                CurrentX += 1
            End While
            NextColumn.Width = CurrentX - NextColumn.Left
            GraphArray.Add(NextColumn)
        Next


        Using g As Graphics = Canvas.CreateGraphics()
            g.Clear(Color.White)
            Dim BlackPen As Pen = New Pen(Brushes.Black)
            Dim RedPen As Pen = New Pen(Brushes.Red)
            Dim VerticalFormat As System.Drawing.StringFormat = New System.Drawing.StringFormat
            Dim LabelFont As Font = dgvDistributionSpreadsheet.DefaultCellStyle.Font
            VerticalFormat.FormatFlags = StringFormatFlags.DirectionVertical

            'Draw the distribution. Also adjusts the rectangles for each column to be exactly the height of the column
            For index = DistMin To DistMax
                Dim NewColumn As Rectangle
                NewColumn.Location = New Point(GraphArray(index).Left, YScale * Distribution(index) + Canvas.Height - ScaleY * 80)
                NewColumn.Width = GraphArray(index).Width
                NewColumn.Height = -YScale * Distribution(index)
                GraphArray(index) = NewColumn
                If index > DistMin Then g.DrawLine(BlackPen, New Point(GraphArray(index - 1).Right, GraphArray(index - 1).Top), NewColumn.Location)
                g.DrawLine(BlackPen, NewColumn.Location, New Point(NewColumn.Right, NewColumn.Top))

            Next

            'Draw the x axis
            g.DrawLine(RedPen, 0, Canvas.Height - ScaleY * 80, Canvas.Width, Canvas.Height - ScaleY * 80)

            'Write labels. Space far enough to not have text overlap
            Dim LastLabelX As Integer = -TextSize("Test", LabelFont, False).Height
            For index = DistMin To DistMax
                If GraphArray(index).Left + GraphArray(index).Width / 2 > LastLabelX + TextSize(XArray(index), LabelFont, False).Height Then
                    g.DrawString(XArray(index), LabelFont, Brushes.Black, New Point(GraphArray(index).Left + GraphArray(index).Width / 2 - TextSize(XArray(index), LabelFont, False).Height / 2, Canvas.Height - ScaleY * 70), VerticalFormat)
                    LastLabelX = GraphArray(index).Left + GraphArray(index).Width / 2
                End If
            Next
        End Using

        Return GraphArray

    End Function

    Private Function DrawInternalRectangle(g As Graphics, OuterRectangle As Rectangle, FillColour As Color) As Boolean

        Try
            'Create a solid pen of the nominated colour
            Dim FillBrush As Brush = New SolidBrush(FillColour)
            Dim FillPen As Pen = New Pen(FillBrush)

            'Define a rectangle 1 pixel smaller in each dimension
            Dim InternalRectangle As Rectangle
            InternalRectangle.X = OuterRectangle.X + 1
            InternalRectangle.Y = OuterRectangle.Y + 1
            InternalRectangle.Width = OuterRectangle.Width - 1
            InternalRectangle.Height = OuterRectangle.Height - 1

            'Draw the internal rectangle with solid colour of the nominated colour
            g.FillRectangle(FillBrush, InternalRectangle)
        Catch
            'Return False if there is an error
            Return False
        End Try

        'Return True if performed without errors
        Return True

    End Function

    Private Sub picDistributionHistogram_MouseDown(sender As Object, e As MouseEventArgs) Handles picDistributionHistogram.MouseDown

        If ColumnArray Is Nothing OrElse ColumnArray.Count = 0 Then Exit Sub

        HistogramSelection = True
        For Each column As Rectangle In SelectedColumns
            DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), column, Color.White)
        Next
        SelectedColumns.Clear()
        For Each Row As DataGridViewRow In dgvDistributionSpreadsheet.Rows
            Row.Selected = False
        Next

        For Each Column As Rectangle In ColumnArray
            If Column.Contains(e.Location) Then
                SelectedColumns.Add(Column)
                DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), Column, Color.Blue)
                dgvDistributionSpreadsheet.Rows(ColumnArray.IndexOf(Column)).Selected = True
                dgvDistributionSpreadsheet.FirstDisplayedScrollingRowIndex = dgvDistributionSpreadsheet.Rows.IndexOf(dgvDistributionSpreadsheet.SelectedRows(0))
            End If
        Next
        HistogramSelection = False

    End Sub

    Private Sub picDistributionHistogram_MouseMove(sender As Object, e As MouseEventArgs) Handles picDistributionHistogram.MouseMove
        If e.Button <> MouseButtons.Left Then Exit Sub

        If SelectedColumns.Count = 0 Then Exit Sub

        HistogramSelection = True

        Dim Direction As Integer = Math.Sign(e.X - SelectedColumns(0).X - SelectedColumns(0).Width / 2)
        If Direction = 0 Then Direction = 1

        Dim RemovedColumns As List(Of Rectangle) = New List(Of Rectangle)

        For Each column As Rectangle In SelectedColumns
            If Direction = 1 Then
                If column.X < SelectedColumns(0).X Then
                    DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), column, Color.White)
                    RemovedColumns.Add(column)

                End If
            Else
                If column.X > SelectedColumns(0).X Then
                    DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), column, Color.White)
                    RemovedColumns.Add(column)
                End If
            End If

        Next
        For Each column As Rectangle In RemovedColumns
            dgvDistributionSpreadsheet.Rows(ColumnArray.IndexOf(column)).Selected = False
            SelectedColumns.Remove(column)
        Next

        For Each column As Rectangle In SelectedColumns
            If e.X * Direction < column.X * Direction Then
                If SelectedColumns.IndexOf(column) = 0 Then Continue For
                DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), column, Color.White)
                RemovedColumns.Add(column)
            End If
        Next
        For Each column As Rectangle In RemovedColumns
            dgvDistributionSpreadsheet.Rows(ColumnArray.IndexOf(column)).Selected = False
            SelectedColumns.Remove(column)
        Next

        For x = SelectedColumns(0).X + SelectedColumns(0).Width / 2 To e.X Step Direction
            For Each column As Rectangle In ColumnArray
                If x >= column.Left And x <= column.Right And SelectedColumns.Contains(column) = False Then
                    DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), column, Color.Blue)
                    SelectedColumns.Add(column)
                    dgvDistributionSpreadsheet.Rows(ColumnArray.IndexOf(column)).Selected = True
                End If
            Next
        Next

        If dgvDistributionSpreadsheet.SelectedRows.Count > 0 Then
            Dim TopRow As Integer = dgvDistributionSpreadsheet.RowCount - 1
            For Each Row As DataGridViewRow In dgvDistributionSpreadsheet.SelectedRows
                If Row.Index < TopRow Then TopRow = Row.Index
            Next
            dgvDistributionSpreadsheet.FirstDisplayedScrollingRowIndex = TopRow
        End If

        HistogramSelection = False

    End Sub

    Private Sub dgvDistributionSpreadsheet_SelectionChanged(sender As Object, e As EventArgs) Handles dgvDistributionSpreadsheet.SelectionChanged
        If HistogramSelection = True OrElse ColumnArray Is Nothing OrElse ColumnArray.Count = 0 Then Exit Sub
        SpreadsheetSelection = True
        SelectedColumns.Clear()
        For Each column As Rectangle In ColumnArray
            If column.Width > 0 Then
                DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), column, Color.White)
            End If
        Next
        For Each row As DataGridViewRow In dgvDistributionSpreadsheet.SelectedRows
            SelectedColumns.Add(ColumnArray(row.Index))
            DrawInternalRectangle(picDistributionHistogram.CreateGraphics(), ColumnArray(row.Index), Color.Blue)
        Next
        SpreadsheetSelection = False
    End Sub
End Class
