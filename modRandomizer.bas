Attribute VB_Name = "modRandomizer"
Option Compare Database

Public Function Randomizer()
    Dim result As Double
    
    Randomize
    result = Rnd()
    
    Randomizer = result
    
End Function
