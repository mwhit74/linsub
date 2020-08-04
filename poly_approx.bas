Attribute VB_Name = "poly_approx"
Option Explicit

Function poly2_coeffs(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double) As Variant
    Dim A(2, 2) As Variant
    Dim b(2) As Variant
    Dim abc_soln As Variant
    
    A(0, 0) = x1 ^ 2
    A(0, 1) = x1
    A(0, 2) = 1
    
    A(1, 0) = x2 ^ 2
    A(1, 1) = x2
    A(1, 2) = 1
    
    A(2, 0) = x3 ^ 2
    A(2, 1) = x3
    A(2, 2) = 1
    
    b(0) = y1
    b(1) = y2
    b(2) = y3
    
    abc_soln = linalg.single_solve(A, b)
    
    poly2_coeffs = abc_soln
    
End Function

Sub abc()
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim x3 As Double
    Dim y3 As Double
    
    Dim arr As Variant
    
    x1 = Worksheets("Sheet2").Range("C4").Value
    y1 = Worksheets("Sheet2").Range("D4").Value
    x2 = Worksheets("Sheet2").Range("E4").Value
    y2 = Worksheets("Sheet2").Range("F4").Value
    x3 = Worksheets("Sheet2").Range("G4").Value
    y3 = Worksheets("Sheet2").Range("H4").Value
    
    arr = poly2_coeffs(x1, y1, x2, y2, x3, y3)
    
    Worksheets("Sheet2").Range("I4").Value = arr(0)
    Worksheets("Sheet2").Range("J4").Value = arr(1)
    Worksheets("Sheet2").Range("K4").Value = arr(2)
    
    
End Sub
