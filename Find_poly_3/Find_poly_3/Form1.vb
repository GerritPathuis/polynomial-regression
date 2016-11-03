Imports System.Text
Imports System
Imports System.Configuration
Imports System.Math

Public Structure PPOINT
    Public x As Double
    Public y As Double
End Structure
'--------------------------------------------------------------------------------------------------
' Polynomial regression In VB.net (Visual Studio 2015 Windows Desktop)
'   SEE http://www.vbforums.com/showthread.php?311225-Polynomials-Resolved&highlight=polynomials
'
'--------------------------------------------------------------------------------------------------

Public Class Form1
    Public PZ1(33) As PPOINT  'Raw data

    Public BQ(,) As Double   'Poly Coefficients
    Dim words() As String

    '-------------- Ventilatoren Schets T1E --------------------------
    Dim TFlow() As Double = {0.00, 2.85, 3.8, 4.28, 4.75, 5.25, 5.7, 6.25, 6.65, 7.6, 8.55, 9.5}
    Dim Tpstat() As Double = {3179.6, 3455.4, 3452.3, 3389.0, 3279.2, 3121.7, 2934.4, 2654.5, 2413.4, 1792.8, 1156.9, 475.0}
    Dim Tptot() As Double = {3179.6, 3509.0, 3555.0, 3510.0, 3432.4, 3306.0, 3148.9, 2912.0, 2704.6, 2175.9, 1639.6, 1072.6}

    '-------------- Cyclone, deeltje grootte verlies cijfers--------- 
    Dim Max_Particle_dia() As Double = {2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 25, 30, 35}
    Dim ac300() As Double = {97.0, 76.0, 54.0, 45.0, 36.0, 29.0, 20.5, 14.0, 11.0, 8.4, 5.5, 4.2, 3.2}

    '-------------- poly results--------------------------
    Dim poly1() As Double = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    Dim poly2() As Double = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim j As Integer
        Dim t() As PPOINT

        'ReDim PZ1(TFlow.Length - 1)                'If PZ too big wrong results will result !!!!!!
        ReDim PZ1(Max_Particle_dia.Length - 1)      'If PZ too big wrong results will result !!!!!!
        '--------------------Ptotal------------------------------------
        For j = 0 To (Max_Particle_dia.Length - 1)
            'PZ1(j).x = TFlow(j)        'Fan data
            'PZ1(j).y = Tptot(j)        'Fan data
            PZ1(j).x = Max_Particle_dia(j) / 2      'Cyclone data
            PZ1(j).y = ac300(j)                     'Cyclone data
        Next j
        t = Trend(PZ1, 5)
        For j = 0 To (TFlow.Length - 1)
            poly1(j) = BQ(0, 0) + BQ(1, 0) * PZ1(j).x ^ 1 + BQ(2, 0) * PZ1(j).x ^ 2 + BQ(3, 0) * PZ1(j).x ^ 3 + BQ(4, 0) * PZ1(j).x ^ 4 + BQ(5, 0) * PZ1(j).x ^ 5
        Next j

        ''-------------------Pstat------------------------------------
        'For j = 0 To (TFlow.Length - 1)
        '    PZ1(j).x = TFlow(j)
        '    PZ1(j).y = Tpstat(j)
        'Next j
        't = Trend(PZ1, 5)
        'For j = 0 To (TFlow.Length - 1)
        '    poly2(j) = BQ(0, 0) + BQ(1, 0) * PZ1(j).x ^ 1 + BQ(2, 0) * PZ1(j).x ^ 2 + BQ(3, 0) * PZ1(j).x ^ 3 + BQ(4, 0) * PZ1(j).x ^ 4 + BQ(5, 0) * PZ1(j).x ^ 5
        'Next j

        '--------------------------------------------------------
        draw_chart1()
    End Sub

    Public Function Trend(Data() As PPOINT, ByVal Degree As Integer) As PPOINT()
        'degree 1 = straight line y=a+bx
        'degree n = polynomials!!


        Dim a(,), Ai(,), MP(,) As Double             '2 Dimensional arrays
        Dim SigmaA(), SigmaP() As Double             '1 Dimensional arrays
        Dim PointCount, MaxTerm, m, n, i, j As Integer
        Dim Ret() As PPOINT
        Dim Equation As String

        Degree = Degree + 1

        MaxTerm = (2 * (Degree - 1))
        PointCount = Data.Length

        ReDim SigmaA(MaxTerm - 1)
        ReDim SigmaP(MaxTerm - 1)

        ' Get the coefficients lists for matrices A, and P
        For m = 0 To (MaxTerm - 1)
            For n = 0 To (PointCount - 1)
                ' MessageBox.Show(Data(n).x)
                SigmaA(m) = SigmaA(m) + (Data(n).x ^ (m + 1))
                SigmaP(m) = SigmaP(m) + ((Data(n).x ^ m) * Data(n).y)
            Next
        Next

        ' Create Matrix A, and fill in the coefficients
        ReDim a(Degree - 1, Degree - 1)

        For i = 0 To (Degree - 1)
            For j = 0 To (Degree - 1)
                If i = 0 And j = 0 Then
                    a(i, j) = PointCount
                Else
                    a(i, j) = SigmaA((i + j) - 1)
                End If
            Next
        Next

        ' Create Matrix MP, and fill in the coefficients
        ReDim MP(Degree - 1, 0)
        For i = 0 To (Degree - 1)
            MP(i, 0) = SigmaP(i)
        Next

        ' We have A, and MP of AB=MP, so we can solve B because B=AiP
        Ai = MxInverse(a)
        BQ = MxMultiplyCV(Ai, MP)

        ' Now we solve the equations and generate the list of points
        PointCount = PointCount - 1
        ReDim Ret(PointCount)

        ' Work out non exponential first term
        For i = 0 To PointCount
            Ret(i).X = Data(i).x
            Ret(i).y = BQ(0, 0)
        Next

        ' Work out other exponential terms including exp 1
        For i = 0 To PointCount
            For j = 1 To Degree - 1
                Ret(i).y = Ret(i).y + (BQ(j, 0) * Ret(i).x ^ j)
            Next
        Next

        '-------- show the coefficients-------------
        Equation = "y=" & Format$(BQ(0, 0), "0.000000000000") & " + "
        For j = 1 To (Degree - 1)
            Equation = Equation & Format$(BQ(j, 0), "0.000000000000") & "x^" & j & " + "
        Next
        Equation = Microsoft.VisualBasic.Left(Equation, Len(Equation) - 3)
        MessageBox.Show(Equation)

        Trend = Ret
    End Function

    Public Function MxMultiplyCV(Matrix1(,) As Double, ColumnVector(,) As Double) As Double(,)

        Dim i, j As Integer
        Dim Rows, Cols As Integer
        Dim Ret(,) As Double        '2 Dimensional array

        Rows = Matrix1.GetLength(0) - 1
        Cols = Matrix1.GetLength(1) - 1

        ReDim Ret(ColumnVector.GetLength(0) - 1, 0) 'returns a column vector

        For i = 0 To Rows
            For j = 0 To Cols
                Ret(i, 0) = Ret(i, 0) + (Matrix1(i, j) * ColumnVector(j, 0))
            Next
        Next

        MxMultiplyCV = Ret
    End Function

    Public Function MxInverse(Matrix(,) As Double) As Double(,)
        Dim i, j As Integer
        Dim Rows, Cols As Integer          '1 Dimensional array
        Dim Tmp(,), Ret(,) As Double    '2 Dimensional array
        Dim Degree As Integer

        Tmp = Matrix

        Rows = Tmp.GetLength(0) - 1     'First dimension of the array
        Cols = Tmp.GetLength(1) - 1     'Second dimension of the array

        Degree = Cols + 1

        'Augment Identity matrix onto matrix M to get [M|I]

        ReDim Preserve Tmp(Rows, (Degree * 2) - 1)
        For i = Degree To (Degree * 2) - 1
            Tmp((i Mod Degree), i) = 1
        Next

        ' Now find the inverse using Gauss-Jordan Elimination which should get us [I|A-1]
        MxGaussJordan(Tmp)

        ' Copy the inverse (A-1) part to array to return
        ReDim Ret(Rows, Cols)
        For i = 0 To Rows
            For j = Degree To (Degree * 2) - 1
                Ret(i, j - Degree) = Tmp(i, j)
            Next
        Next

        MxInverse = Ret
    End Function

    Public Sub MxGaussJordan(Matrix(,) As Double)

        Dim Rows As Integer
        Dim Cols As Integer
        Dim P As Integer
        Dim i As Integer
        Dim j As Integer
        Dim m As Double
        Dim d As Double
        Dim Pivot As Double

        Rows = Matrix.GetLength(0) - 1     'First dimension of the array
        Cols = Matrix.GetLength(1) - 1     'Second dimension of the array

        ' Reduce so we get the leading diagonal
        For P = 0 To Rows
            Pivot = Matrix(P, P)
            For i = 0 To Rows
                If Not P = i Then
                    m = Matrix(i, P) / Pivot
                    For j = 0 To Cols
                        Matrix(i, j) = Matrix(i, j) + (Matrix(P, j) * -m)
                    Next
                End If
            Next
        Next

        'Divide through to get the identity matrix
        'Note: the identity matrix may have very small values (close to zero)
        'because of the way floating points are stored.
        For i = 0 To Rows
            d = Matrix(i, i)
            For j = 0 To Cols
                Matrix(i, j) = Matrix(i, j) / d
            Next
        Next
    End Sub

    Private Sub draw_chart1()
        Dim hh As Integer

        Try
            'Clear all series And chart areas so we can re-add them
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()

            Chart1.Series.Add("Series0")    'Input
            Chart1.Series.Add("Series1")    'Poly line
            Chart1.Series.Add("Series2")    'Poly line

            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            Chart1.Series(1).ChartArea = "ChartArea0"
            Chart1.Series(2).ChartArea = "ChartArea0"

            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series(2).ChartType = DataVisualization.Charting.SeriesChartType.Line

            Chart1.Titles.Add("Polynomial regression")
            Chart1.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

            Chart1.Series(0).Name = "Input"
            Chart1.Series(1).Name = "Poly1"
            Chart1.Series(2).Name = "Poly2"

            Chart1.Series(0).Color = Color.Blue
            Chart1.Series(1).Color = Color.Red
            Chart1.Series(2).Color = Color.Yellow

            '----------- labels on-off ------------------
            Chart1.Series(0).IsValueShownAsLabel = True
            Chart1.Series(1).IsValueShownAsLabel = True

            Chart1.Series(0).BorderWidth = 1
            Chart1.Series(1).BorderWidth = 1
            Chart1.Series(2).BorderWidth = 1

            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True

            '-------------------Chart data ---------------------
            For hh = 0 To (TFlow.Length - 1)
                'Chart1.Series(0).Points.AddXY(PZ1(hh).x, PZ1(hh).y)           'Present the Raw data
                Chart1.Series(1).Points.AddXY(PZ1(hh).x, poly1(hh))
                Chart1.Series(2).Points.AddXY(PZ1(hh).x, poly2(hh))
            Next hh
            'Chart1.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub
End Class
