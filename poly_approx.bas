Attribute VB_Name = "poly_approx"
Option Explicit

Function poly2_coeffs(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double) As Variant
    Dim a(2, 2) As Variant
    Dim b(2) As Variant
    Dim Aov() As Variant
    Dim lu() As Variant
    Dim ov() As Variant
    Dim abc_soln As Variant
    Dim lus As String
    
    a(0, 0) = x1 ^ 2
    a(0, 1) = x1
    a(0, 2) = 1
    
    a(1, 0) = x2 ^ 2
    a(1, 1) = x2
    a(1, 2) = 1
    
    a(2, 0) = x3 ^ 2
    a(2, 1) = x3
    a(2, 2) = 1
    
    b(0) = y1
    b(1) = y2
    b(2) = y3
    
    Aov = linsub.lu_decomp(a)
    lu = Aov(0)
    ov = Aov(1)
     
    abc_soln = linsub.solve(lu, ov, b)
    
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
    
    x1 = Worksheets("Sheet1").Range("C4").Value
    y1 = Worksheets("Sheet1").Range("D4").Value
    x2 = Worksheets("Sheet1").Range("E4").Value
    y2 = Worksheets("Sheet1").Range("F4").Value
    x3 = Worksheets("Sheet1").Range("G4").Value
    y3 = Worksheets("Sheet1").Range("H4").Value
    
    arr = poly2_coeffs(x1, y1, x2, y2, x3, y3)
    
    Worksheets("Sheet1").Range("J4").Value = arr(0)
    Worksheets("Sheet1").Range("K4").Value = arr(1)
    Worksheets("Sheet1").Range("L4").Value = arr(2)
    
    
End Sub
