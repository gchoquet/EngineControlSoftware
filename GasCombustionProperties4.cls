VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GasCombustionProperties4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class GasCombustionProperties
' Purpose: Calculate combustion properties of natural gas fuels including higher heating value,
' lower heating value, specific gravity, stoichiometric air/fuel ratio,and methane number.
' Developed by Gary Choquette, Optimized Technical Solutions, LLC
'
' Copyright 2014, Optimized Technical Solutions, LLC

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, version 3 of the License.
'    http://opensource.org/licenses/GPL-3.0

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
  
'Should not be used where H2 is > 1%


Option Explicit

  Private dblCarbonDioxide As Double  'in mole fraction
  Private dblHHV As Double
  Private dblHydrogenSulfide As Double  'in mole fraction
  Private dblNitrogen As Double   'in mole fraction
  Private dblSPGrav As Double
  Private isDirty As Boolean
  
  Private dblMNRaw As Double
  Private dblCO2Adj As Double
  Private dblH2SAdj As Double
  Private dblMethaneNumber As Double
    
  Public Property Let CarbonDioxide(ByVal value As Double)
    If Not value = dblCarbonDioxide Then
      isDirty = True
      dblCarbonDioxide = value
    End If
  End Property
  
  Public Property Get CarbonDioxide() As Double
     CarbonDioxide = dblCarbonDioxide
  End Property

    Public Property Let HigherHeatingValue(ByVal value As Double)
    If Not dblHHV = value Then
      dblHHV = value
      isDirty = True
    End If
  End Property
  
  Public Property Get HigherHeatingValue() As Double
    HigherHeatingValue = dblHHV
  End Property
  
  Public Property Let HydrogenSulfide(ByVal value As Double)
    If Not value = dblHydrogenSulfide Then
      isDirty = True
      dblHydrogenSulfide = value
    End If
  End Property
  
  Public Property Get HydrogenSulfide() As Double
     HydrogenSulfide = dblHydrogenSulfide
  End Property

  Public Property Get MethaneNumber() As Double
    If isDirty Then
      CalculateProperties
    End If
    MethaneNumber = dblMethaneNumber
  End Property

  Public Property Let Nitrogen(ByVal value As Double)
    If Not value = dblNitrogen Then
      isDirty = True
      dblNitrogen = value
    End If
  End Property
  
  Public Property Get Nitrogen() As Double
     Nitrogen = dblNitrogen
  End Property

  Public Property Let SpecificGravity(ByVal value As Double)
    If Not dblSPGrav = value Then
      dblSPGrav = value
      isDirty = True
    End If
  End Property
  
  Public Property Get SpecificGravity() As Double
    SpecificGravity = dblSPGrav
  End Property
  
  Private Sub CalculateProperties()
    Dim hcrAdj As Double
    Dim sgAdj As Double
    sgAdj = (dblSPGrav * 28.964 - dblCarbonDioxide * 44.01 - dblNitrogen * 28.013) / 28.964 / (1 - dblCarbonDioxide - dblNitrogen)
    hcrAdj = 1.800219371 + 1.216463553 / sgAdj
    dblMNRaw = 34.26085951 * hcrAdj + 2.944058733E-09 * sgAdj * (21.30716314 * hcrAdj ^ 2) ^ hcrAdj - 59.1854841
    CalcCO2Adjustment
    CalcH2SAdjustment
    dblMethaneNumber = dblMNRaw + dblCO2Adj + dblH2SAdj
    isDirty = False
  End Sub

  Private Sub CalcH2SAdjustment()
    If dblHydrogenSulfide <> 0 Then
      dblH2SAdj = -14.83 * dblHydrogenSulfide - 1.048 * dblHydrogenSulfide ^ 0.002036
    Else
      dblH2SAdj = 0
    End If
  End Sub

  Private Sub CalcCO2Adjustment()
    Dim mnAdj As Double
    If dblCarbonDioxide <= 0 Then
      dblCO2Adj = 0
    Else
      If dblMNRaw < 70 Then
        mnAdj = 70
      Else
        mnAdj = dblMNRaw
      End If
      dblCO2Adj = 57.0952 * dblCarbonDioxide ^ 2 * mnAdj + 0.0000011611 * dblCarbonDioxide * mnAdj ^ 4 - 0.081104 - 45.2077 * dblCarbonDioxide - 2153.3903 * dblCarbonDioxide ^ 2 - 253.9675 * dblCarbonDioxide ^ 3 - 0.337730332431225 * dblCarbonDioxide ^ 2 * mnAdj ^ 2
    End If
  End Sub

