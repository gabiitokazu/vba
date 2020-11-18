Sub LotkaVolterra()

'------------------------------------------------------------------
'Lab 2 - BSEN 6250 - Aug 27, 2020
'Ana Gabriela Itokazu Canzian da Silva

'This program was developed to simulate predator-prey populations relationship
'considering the Lotka Voltera Model of predation
'Utilizing initial variables provided on the spreadsheet and
'displaying graphical representations for that simulation
'developing outputs for both the model and descriptive figures, like min and max values
'------------------------------------------------------------------

Dim r, g, m, h As Double 'rates
Dim N, P As Double 'Predator and prey Populations
Dim Duration, dt As Double 'allowing the user to use decimals, to enhance definition of the model if desired

r = Range("Inputs").Cells(1).Value 'growth rate at min pop pre, 1/time
g = Range("Inputs").Cells(2).Value 'predation efficiency, how many preys die per each predator
m = Range("Inputs").Cells(3).Value 'predator death rate as a function of prey pop, how the scarcicity of food affects predators 1.time^(-1)
h = Range("Inputs").Cells(4).Value 'capture rate, conversion prey into predators,1.#prey^(-1)time^(-1)
N = Range("Inputs").Cells(5).Value 'prey population, #prey
P = Range("Inputs").Cells(6).Value 'predator population, #pred
Duration = Range("Inputs").Cells(7).Value 'total duration of the simulation, years
dt = Range("Inputs").Cells(8).Value 'definition of the model, timestep range, years

Call ClearContents

'Get initial values (first row output) from inputs + initial time cell
Row = 1
Range("Output").Cells(Row, 1).Value = 0
Range("Output").Cells(Row, 2).Value = N
Range("Output").Cells(Row, 3).Value = P

'set initial conditions

LastP = P
LastN = N

For t = dt To Duration Step dt
    Row = Row + 1               'Increment row number for output
    'N = LastN + (r - g * LastP) * LastN * dt
    'P = LastP + (h - m * LastN) * LastP * dt
    
    N = LastN + ((r * LastN) - (g * LastN * LastP)) * dt
    P = LastP + ((h * LastN * LastP) - (m * LastP)) * dt
    
    LastN = N
    LastP = P

    Range("Output").Cells(Row, 1).Value = t
    Range("Output").Cells(Row, 2).Value = N 'Write N on output table
    Range("Output").Cells(Row, 3).Value = P 'Write P on output table
Next t

Call MaxMin


End Sub

Sub MaxMin()

Dim MinPrey, MaxPrey, MinPred, MaxPred As Double
    
Set Prey = Worksheets("LotkaVolterra").Range("PreySim")
Set Pred = Worksheets("LotkaVolterra").Range("PredSim")

MinPrey = Application.WorksheetFunction.Min(Prey)
MinPred = Application.WorksheetFunction.Min(Pred)
MaxPrey = Application.WorksheetFunction.Max(Prey)
MaxPred = Application.WorksheetFunction.Max(Pred)

Range("Output2").Cells(1, 1).Value = MinPrey
Range("Output2").Cells(2, 1).Value = MinPred
Range("Output2").Cells(1, 2).Value = MaxPrey
Range("Output2").Cells(2, 2).Value = MaxPred

End Sub


' Routine to clear the contents from the "Output" and "Output2" section

Sub ClearContents()
    
    Range("Output").ClearContents
    Range("Output2").ClearContents

End Sub


'--------------------------------------------------------------------
' Program for the cP simulation
'--------------------------------------------------------------------

Sub cP()

Dim r, g, m, h As Double 'rates
Dim N, P As Double 'predator and prey populations
Dim Duration As Double 'allowing the user to use decimals for duration
Dim dt As Double 'allowing the user to use decimals for timestep, to enhance definition of the model if desired
Dim cP As Double '

r = Range("Inputs_cP").Cells(1).Value 'growth rate at min pop pre, 1/time
g = Range("Inputs_cP").Cells(2).Value 'predation efficiency, how many preys die per each predator
m = Range("Inputs_cP").Cells(3).Value 'predator death rate as a function of prey pop, how the scarcicity of food affects predators 1.time^(-1)
h = Range("Inputs_cP").Cells(4).Value 'capture rate, conversion prey into predators,1.#prey^(-1)time^(-1)
N = Range("Inputs_cP").Cells(5).Value 'prey population, #prey
P = Range("Inputs_cP").Cells(6).Value 'predator population, #pred
Duration = Range("Inputs_cP").Cells(7).Value 'total duration of the simulation, years
dt = Range("Inputs_cP").Cells(8).Value 'definition of the model, timestep range, years
c = Range("Inputs_cP").Cells(9).Value 'harvesting term

Call ClearContents_cP

'Get initial values (first row output) from inputs + initial time cell
Row = 1
Range("Output_cP").Cells(Row, 1).Value = 0
Range("Output_cP").Cells(Row, 2).Value = N
Range("Output_cP").Cells(Row, 3).Value = P

'set initial conditions

LastP = P
LastN = N

For t = dt To Duration Step dt
    Row = Row + 1               'Increment row number for output

    
    N = LastN + ((r * LastN) - (g * LastN * LastP)) * dt
    P = LastP + ((h * LastN * LastP) - (m * LastP) - (c * LastP)) * dt
    
    LastN = N
    LastP = P

    Range("Output_cP").Cells(Row, 1).Value = t
    Range("Output_cP").Cells(Row, 2).Value = N 'Write N on output table
    Range("Output_cP").Cells(Row, 3).Value = P 'Write P on output table
Next t

Call MaxMin_cP


End Sub

Sub MaxMin_cP()

Dim MinPrey, MaxPrey, MinPred, MaxPred As Double
    
Set Prey = Worksheets("cP").Range("PreySim_cP")
Set Pred = Worksheets("cP").Range("PredSim_cP")

MinPrey = Application.WorksheetFunction.Min(Prey)
MinPred = Application.WorksheetFunction.Min(Pred)
MaxPrey = Application.WorksheetFunction.Max(Prey)
MaxPred = Application.WorksheetFunction.Max(Pred)

Range("Output2_cP").Cells(1, 1).Value = MinPrey
Range("Output2_cP").Cells(2, 1).Value = MinPred
Range("Output2_cP").Cells(1, 2).Value = MaxPrey
Range("Output2_cP").Cells(2, 2).Value = MaxPred

End Sub


' Routine to clear the contents from the output sections

Sub ClearContents_cP()
    
    Range("Output_cP").ClearContents
    Range("Output2_cP").ClearContents

End Sub

'---------------------------------------------------------------------------
'End of Program for cP simulation
'---------------------------------------------------------------------------




'--------------------------------------------------------------------
' Program for the K simulation
'--------------------------------------------------------------------

Sub K()

Dim r, g, m, h As Double 'rates
Dim N, P As Double 'predator and prey populations
Dim Duration As Double 'allowing the user to use decimals for duration
Dim dt As Double 'allowing the user to use decimals for timestep, to enhance definition of the model if desired
Dim K As Double '

r = Range("Inputs_K").Cells(1).Value 'growth rate at min pop pre, 1/time
g = Range("Inputs_K").Cells(2).Value 'predation efficiency, how many preys die per each predator
m = Range("Inputs_K").Cells(3).Value 'predator death rate as a function of prey pop, how the scarcicity of food affects predators 1.time^(-1)
h = Range("Inputs_K").Cells(4).Value 'capture rate, conversion prey into predators,1.#prey^(-1)time^(-1)
N = Range("Inputs_K").Cells(5).Value 'prey population, #prey
P = Range("Inputs_K").Cells(6).Value 'predator population, #pred
Duration = Range("Inputs_K").Cells(7).Value 'total duration of the simulation, years
dt = Range("Inputs_K").Cells(8).Value 'definition of the model, timestep range, years
K = Range("Inputs_K").Cells(9).Value 'carrying capacity term

Call ClearContents_K

'Get initial values (first row output) from inputs + initial time cell
Row = 1
Range("Output_K").Cells(Row, 1).Value = 0
Range("Output_K").Cells(Row, 2).Value = N
Range("Output_K").Cells(Row, 3).Value = P

'set initial conditions

LastP = P
LastN = N

For t = dt To Duration Step dt
    Row = Row + 1               'Increment row number for output

    N = LastN + ((r * LastN * (1 - (LastN / K))) - (g * LastN * LastP)) * dt
    P = LastP + ((h * LastN * LastP) - (m * LastP)) * dt
    
    LastN = N
    LastP = P

    Range("Output_K").Cells(Row, 1).Value = t
    Range("Output_K").Cells(Row, 2).Value = N 'Write N on output table
    Range("Output_K").Cells(Row, 3).Value = P 'Write P on output table
Next t

Call MaxMin_K


End Sub


Sub MaxMin_K()

Dim MinPrey, MaxPrey, MinPred, MaxPred As Double
    
Set Prey = Worksheets("K").Range("PreySim_K")
Set Pred = Worksheets("K").Range("PredSim_K")

MinPrey = Application.WorksheetFunction.Min(Prey)
MinPred = Application.WorksheetFunction.Min(Pred)
MaxPrey = Application.WorksheetFunction.Max(Prey)
MaxPred = Application.WorksheetFunction.Max(Pred)

Range("Output2_K").Cells(1, 1).Value = MinPrey
Range("Output2_K").Cells(2, 1).Value = MinPred
Range("Output2_K").Cells(1, 2).Value = MaxPrey
Range("Output2_K").Cells(2, 2).Value = MaxPred

End Sub


' Routine to clear the contents from the output sections

Sub ClearContents_K()
    
    Range("Output_K").ClearContents
    Range("Output2_K").ClearContents

End Sub


'---------------------------------------------------------------------------
'End of Program for K simulation
'---------------------------------------------------------------------------

