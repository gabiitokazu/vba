Sub deer()

'This program computes population dynamics for a deer hurd based on
'equations outlined in section 9.5 of the book "Computer Simulation
'in Biology: A BASIC introduction" by Keen and Spain.


'-------------------------------------------------------------------
'       Model Inputs
'-------------------------------------------------------------------
Dim s(100) As Double    'Fraction of individuals surviving age class
Dim m(100) As Double    'reproduction rate for age class
Dim Females(12) As Double   'Number of females in age class
Dim Males(12) As Double   'Number of males in age class
Dim P(100) As Double    'Probability of mortality due to hunting for age class
Dim hunter As Double    'Hunter density

'-------------------------------------------------------------------
'       Model Variables
'-------------------------------------------------------------------
Dim HuntLoss As Double      'Total males harvested due to hunting
Dim Mature_Males As Double  'Number of mature males in the population
Dim Mature_Females As Double    'Number of mature females in the population
Dim Births As Double        'Births during timestep
Dim Repro_Success As Double 'Reproductive success
Dim Time As Integer         'Time
Dim Tot_Females As Double       'Total number of females
Dim Tot_Males As Double       'Total number of males
Dim Fertility As Double     'Variable to increase fertility rate if males are reduced
Dim CumHuntLoss As Double   'Cumulative deer harvested by hunting
Dim CCAP As Double          'Carrying Capacity
Dim survivors As Double
Dim NatDeath As Double
Dim dMales As Double
Dim dFemales As Double
Dim oldMales(12) As Double
Dim oldFemales(12) As Double


'-------------------------------------------------------------------
'       Read Inputs
'-------------------------------------------------------------------

hunter = Range("C6").Value
CCAP = Range("D4").Value

For i = 0 To 11
    s(i) = Range("Inputs").Cells(i + 1, 2).Value
    m(i) = Range("Inputs").Cells(i + 1, 3).Value
    oldFemales(i) = Range("Inputs").Cells(i + 1, 4).Value
    oldMales(i) = Range("Inputs").Cells(i + 1, 5).Value
    P(i) = Range("Inputs").Cells(i + 1, 6).Value
Next i

'  Clear contents from existing model runs in spreadsheet
Range("Population").ClearContents
Range("Summary").ClearContents
Range("Time").ClearContents

'-------------------------------------------------------------------
'       Main Loop for Simulation
'-------------------------------------------------------------------

For Time = 1 To 100  'Run 100 years on yearly time step

    '---------------------------------------------------------
    ' Compute Number of mature males and females available for reproduciton
    '---------------------------------------------------------
    Mature_Males = 0
    Mature_Females = 0
    For i = 1 To 11
        Mature_Males = Mature_Males + oldMales(i)
        Mature_Females = Mature_Females + oldFemales(i)
    Next i
    
    '---------------------------------------------------------
    ' Compute Birth Rate
    '---------------------------------------------------------
    ' Compute rate of production of fawns for all age classes (sum(mxFx)):
    Repro_Success = 0 '(for i = 0)
    
    For i = 1 To 11
        Repro_Success = Repro_Success + (m(i) * oldFemales(i))
    Next i
    
    ' fraction of females fertilized each year (R):
    
    Fertility = 1 - Exp((-0.002656) * Mature_Males)
    
    ' Total birth each year:    deltaN = Repro_Success * Fertility
    Tot_Males = 0
    Tot_Females = 0
    
    For i = 0 To 11
        Tot_Males = Tot_Males + oldMales(i)
        Tot_Females = Tot_Females + oldFemales(i)
    Next i
    
    deltaN = CCAP - (1.5 * ((Mature_Males + Mature_Females) / 6000))
    
    'Total Births
    Births = Repro_Success * Fertility * deltaN
    
    Births = Births - Births * (1 - s(0))
    
    If Births < 0 Then Births = 0
    
    'For males age 0 (no hunting pressure):

     Males(0) = 0.528 * Births
     
    If Males(0) < 0 Then Males(0) = 0

    Females(0) = 0.472 * Births
    If Females(0) < 0 Then Females(0) = 0
   
  CumHuntLoss = 0
   
   For i = 1 To 11
      HuntLoss = (oldMales(i) * hunter * P(i)) / (hunter + (Tot_Males - oldMales(0)))
      CumHuntLoss = CumHuntLoss + HuntLoss
      
    'governing equation males age 1+:
      survivorsM = (s(i - 1) * oldMales(i - 1))
      NatDeathM = oldMales(i) * (1 - s(i))
      goneM = oldMales(i) * s(i)
      dMales = survivorsM - NatDeathM - goneM - HuntLoss
      Males(i) = oldMales(i) + dMales
      If Males(i) < 0 Then Males(i) = 0
      
      Tot_Males = Tot_Males + Males(i) * dt
      
      'governing equationfemales age 1+:
      survivorsF = (s(i - 1) * oldFemales(i - 1))
      NatDeathF = oldFemales(i) * (1 - s(i))
      goneF = oldFemales(i) * s(i)
      dFemales = survivorsF - NatDeathF - goneF
      Females(i) = oldFemales(i) + dFemales
      If Females(i) < 0 Then Females(i) = 0
      
      Tot_Females = Tot_Females + Females(i) * dt
      
   Next i
   
   Tot_Males = Tot_Males + Males(0)
   Tot_Females = Tot_Females + Females(0)
    
    '---------------------------------------------------------
    '   Write output to screen
    ' --------------------------------------------------------
    For i = 1 To 12
       Range("Population").Cells(Time, i).Value = Males(i - 1) + Females(i - 1)
    Next i
    Range("Time").Cells(Time, 1).Value = Time
    Range("Summary").Cells(Time, 1).Value = Tot_Females
    Range("Summary").Cells(Time, 2).Value = Tot_Males
    Range("Summary").Cells(Time, 3).Value = Tot_Males + Tot_Females
    Range("Summary").Cells(Time, 4).Value = Births
    Range("Summary").Cells(Time, 5).Value = CumHuntLoss
    
    For i = 0 To 11
    
        oldMales(i) = Males(i)
        oldFemales(i) = Females(i)
    
    Next i
    
Next Time  'End of big time loop

End Sub



Sub hunter()


End Sub



Sub K()

    For i = 1 To 12 Step 1
        Range("D4").Value = Range("input_k").Cells(i).Value
        Call deer
        Range("output_k").Cells(i, 1).Value = Range("Summary").Cells(100, 3)
    Next i



End Sub

End Sub
