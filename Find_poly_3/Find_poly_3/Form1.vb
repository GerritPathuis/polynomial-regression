Imports System.Text
Imports System
Imports System.Configuration
Imports System.Math

Public Structure POINT
    Public x As Double
    Public y As Double
End Structure
'--------------------------------------------------------------------------------------------------
' Polynomial regression In VB.net (Visual Studio 2015 Windows Desktop)
'   SEE http :  //www.vbforums.com/showthread.php?311225-Polynomials-Resolved&highlight=polynomials
'
'--------------------------------------------------------------------------------------------------

Public Class Form1
    Public P(33) As POINT   'Raw data
    Public B(,) As Double   'Poly Coefficients

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim j As Integer
        Dim t() As POINT

        For j = 0 To 32
            P(j).x = j
            P(j).y = CInt(Math.Ceiling(Rnd() * 50)) + 1
        Next
        t = Trend(P, 5)
        draw_chart1()
    End Sub

    Public Function Trend(Data() As POINT, ByVal Degree As Long) As POINT()
        'degree 1 = straight line y=a+bx
        'degree n = polynomials!!

        Dim a(,) As Double
        Dim Ai(,) As Double

        Dim P(,) As Double
        Dim SigmaA() As Double
        Dim SigmaP() As Double
        Dim PointCount As Long
        Dim MaxTerm As Long
        Dim m As Long, n As Long
        Dim i As Long, j As Long
        Dim Ret() As POINT
        Dim Equation As String

        Degree = Degree + 1

        MaxTerm = (2 * (Degree - 1))
        PointCount = UBound(Data) + 1

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

        ' Create Matrix P, and fill in the coefficients
        ReDim P(Degree - 1, 0)
        For i = 0 To (Degree - 1)
            P(i, 0) = SigmaP(i)
        Next

        ' We have A, and P of AB=P, so we can solve B because B=AiP
        Ai = MxInverse(a)
        B = MxMultiplyCV(Ai, P)

        ' Now we solve the equations and generate the list of points
        PointCount = PointCount - 1
        ReDim Ret(PointCount)

        ' Work out non exponential first term
        For i = 0 To PointCount
            Ret(i).X = Data(i).X
            Ret(i).Y = B(0, 0)
        Next

        ' Work out other exponential terms including exp 1
        For i = 0 To PointCount
            For j = 1 To Degree - 1
                Ret(i).Y = Ret(i).Y + (B(j, 0) * Ret(i).X ^ j)
            Next
        Next

        '-------- show the coefficients-------------
        Equation = "y=" & Format$(B(0, 0), "0.00000") & " + "
        For j = 1 To Degree - 1
            Equation = Equation & Format$(B(j, 0), "0.00000") & "x^" & j & " + "
        Next
        Equation = Microsoft.VisualBasic.Left(Equation, Len(Equation) - 3)
        MessageBox.Show(Equation)

        Trend = Ret
    End Function

    Public Function MxMultiplyCV(Matrix1(,) As Double, ColumnVector(,) As Double) As Double(,)

        Dim i As Long
        Dim j As Long
        Dim Rows As Long
        Dim Cols As Long
        Dim Ret(,) As Double

        Rows = UBound(Matrix1, 1)
        Cols = UBound(Matrix1, 2)

        ReDim Ret(UBound(ColumnVector, 1), 0) 'returns a column vector

        For i = 0 To Rows
            For j = 0 To Cols
                Ret(i, 0) = Ret(i, 0) + (Matrix1(i, j) * ColumnVector(j, 0))
            Next
        Next

        MxMultiplyCV = Ret
    End Function

    Public Function MxInverse(Matrix(,) As Double) As Double(,)
        Dim i As Long
        Dim j As Long
        Dim Rows As Long
        Dim Cols As Long
        Dim Tmp(,) As Double
        Dim Ret(,) As Double
        Dim Degree As Long

        Tmp = Matrix

        Rows = UBound(Tmp, 1)
        Cols = UBound(Tmp, 2)
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

        Dim Rows As Long
        Dim Cols As Long
        Dim P As Long
        Dim i As Long
        Dim j As Long
        Dim m As Double
        Dim d As Double
        Dim Pivot As Double

        Rows = UBound(Matrix, 1)
        Cols = UBound(Matrix, 2)

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
        Dim poly_result As Double

        Try
            'Clear all series And chart areas so we can re-add them
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()

            Chart1.Series.Add("Series0")    'Input
            Chart1.Series.Add("Series1")    'Poly line

            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            Chart1.Series(1).ChartArea = "ChartArea0"

            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line

            Chart1.Titles.Add("Polynomial regression")
            Chart1.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

            Chart1.Series(0).Name = "Input"
            Chart1.Series(1).Name = "Poly"

            Chart1.Series(0).Color = Color.Blue
            Chart1.Series(1).Color = Color.Red

            '----------- labels on-off ------------------
            Chart1.Series(0).IsValueShownAsLabel = True

            Chart1.Series(0).BorderWidth = 4
            Chart1.Series(1).BorderWidth = 3

            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart1.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
            Chart1.ChartAreas("ChartArea0").AxisY2.MinorTickMark.Enabled = True

            Chart1.ChartAreas("ChartArea0").AxisX.Title = "Debiet [Am3/s]"
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Ptotaal [Pa]"
            Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical

            '-------------------Chart data ---------------------
            For hh = 0 To 32
                Chart1.Series(0).Points.AddXY(P(hh).x, P(hh).y)

                poly_result = B(0, 0) + B(1, 0) * P(hh).x ^ 1 + B(2, 0) * P(hh).x ^ 2 + B(3, 0) * P(hh).x ^ 3 + B(4, 0) * P(hh).x ^ 4 + +B(5, 0) * P(hh).x ^ 5
                Chart1.Series(1).Points.AddXY(P(hh).x, poly_result)
            Next hh
            Chart1.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub
End Class
