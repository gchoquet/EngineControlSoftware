Attribute VB_Name = "GasCompCalcs"
Option Explicit

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


' EXCEL VBA interface to classes to estimate methane number and other gas composition properties


Private oMNGEC As New GasCombustionProperties
Private oMNGEC4 As New GasCombustionProperties4

  Public Function MN_Ext(C1 As Double, Optional C2 As Double = 0, Optional C3 As Double = 0, Optional IC4 As Double = 0, Optional NC4 As Double = 0, Optional IC5 As Double = 0, Optional NC5 As Double = 0, Optional C6 As Double = 0, Optional C7 As Double = 0, Optional C8 As Double = 0, Optional C9 As Double = 0, Optional N2 As Double = 0, Optional CO2 As Double = 0, Optional He As Double = 0, Optional CO As Double = 0, Optional H2 As Double = 0, Optional H2S As Double = 0, Optional CompName As String = "") As Variant
    With oMNGEC
      .CompositionType = molePercent
      .CompositionName = CompName
      .Methane = C1 * 100
      .Ethane = C2 * 100
      .Propane = C3 * 100
      .IsoButane = IC4 * 100
      .NormalButane = NC4 * 100
      .IsoPentane = IC5 * 100
      .NormalPentane = NC5 * 100
      .Hexane = C6 * 100
      .Heptane = C7 * 100
      .Octane = C8 * 100
      .Nonane = C9 * 100
      .Nitrogen = N2 * 100
      .CarbonDioxide = CO2 * 100
      .Helium = He * 100
      .CarbonMonoxide = CO * 100
      .Hydrogen = H2 * 100
      .HydrogenSulfide = H2S * 100
      DoEvents
      Dim rtn(3) As Variant
      Dim x(3) As Double
      x(0) = .MethaneNumber
      x(1) = .CO2Adjustment
      x(2) = .H2SAdjustment
      rtn(0) = x(0)
      rtn(1) = x(1)
      rtn(2) = x(2)
      rtn(3) = x(3)
      MN_Ext = rtn
    End With
  End Function
  
  
  Public Function MN_4(HHV As Double, SG As Double, Optional CO2 As Double = 0, Optional N2 As Double = 0, Optional H2S As Double = 0, Optional CompName As String = "") As Double
    With oMNGEC4
      .HigherHeatingValue = HHV
      .HydrogenSulfide = H2S
      .CarbonDioxide = CO2
      .Nitrogen = N2
      .SpecificGravity = SG
      MN_4 = .MethaneNumber
    End With
  End Function


