Sub Lab1_concentration()

Dim Ci As Double        'Concentration inside of cell, mol/liter
Dim Co As Double       'Concentration outsiede of cell, mol/liter
Dim k As Double         'Diffusion coefficient, 1/minute
Dim dt As Double       'Timestep, minutes
Dim C As Double         'Concentration inside cell
Dim row As Integer     'Row that output will be written to
Dim t As Double         'Time, minutes

k = Range("B6").Value
Ci = Range("B7").Value
Co = Range("B8").Value
dt = Range("B9").Value
FinalTime = Range("B10").Value
Range("Computed").ClearContents

'Write the initial concentration at time t=0 into row 1 of
'array called "Computed" that was defined in excel

row = 1
Range("Computed").Cells(row).Value = Ci

For t = dt To FinalTime Step dt                  'Main loop for solving Equation 4
row = row + 1                                           'Increment row number for output
Ci = Ci - k * (Ci - Co) * dt                         'Equation 4
Range("Computed").Cells(row).Value = Ci  'Write Ci to spreadsheet
Next t


End Sub