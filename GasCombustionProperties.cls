VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GasCombustionProperties"
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
  
  Public Enum GasCompositionType
    moleFraction
    molePercent
  End Enum
  
  Public Enum GasCompositionErrors
    noErrors = 0
    compositionZero = 1
    compositionSumLow = 2
    compositionSumHigh = 4
    allCarbon = 8
    allHydrogen = 16
    allInerts = 32
  End Enum
  
  
  
  Private gctCompositionType As GasCompositionType
  
  Private gceErrors As GasCompositionErrors
  
  Private strCompositionName As String
  Private strErrorNotes As String
  
  'normalized mole fractions
  Private dblH2 As Double
  Private dblH2S As Double
  Private dblCO As Double
  Private dblCO2 As Double
  Private dblN2 As Double
  Private dblO2 As Double
  Private dblHE As Double
  Private dblC1 As Double
  Private dblC2 As Double
  Private dblC3 As Double
  Private dblIC4 As Double
  Private dblNC4 As Double
  Private dblIC5 As Double
  Private dblNC5 As Double
  Private dblNEC5 As Double
  Private dblC6 As Double
  Private dblC7 As Double
  Private dblC8 As Double
  Private dblC9 As Double
  Private dblH2O As Double
  
  
  Private dblBasePressue As Double
  
  'raw values
  Private dblCarbonDioxide As Double
  Private dblCarbonMonoxide As Double
  Private dblEthane As Double
  Private dblNitrogen As Double
  Private dblHelium As Double
  Private dblHeptane As Double
  Private dblHexane As Double
  Private dblHydrogen As Double
  Private dblHydrogenSulfide As Double
  Private dblIsoButane As Double
  Private dblIsoPentane As Double
  Private dblMethane As Double
  Private dblNeoPentane As Double
  Private dblNonane As Double
  Private dblNormalButane As Double
  Private dblNormalPentane As Double
  Private dblOctane As Double
  Private dblOxygen As Double
  Private dblPropane As Double
  Private dblWater As Double

  Private isDirty As Boolean

  'Outputs
  Private dblInerts As Double
  Private dblHHV As Double 'lower heating value in btu/scf
  Private dblLHV As Double 'higher heating value in btu/scf
  Private dblStoichAFR As Double 'stoich AF by volume
  Private dblSPGrav As Double 'specific gravity
  Private dblMSTOICH As Double 'stoich af by mass
  Private dblMethaneNumber As Double
  Private dblH2SAdj As Double
  Private dblCO2Adj As Double
  Private dblWobbeIndex As Double 'wobbe index

  'locals
  Private dblMNRaw As Double
  Private carbonDioxideAirAdjustment As Double
  Private nitrogenAdjustmentFactor As Double
  
  'percent of combustables
  Private dblCMethane As Double
  Private dblCEthane As Double
  Private dblCPropane As Double
  Private dblCPropylene As Double
  Private dblCIsoButane As Double
  Private dblCNormalButane As Double
  Private dblCIsoPentane As Double
  Private dblCNormalPentane As Double
  Private dblCNeoPentane As Double
  Private dblCHexane As Double
  Private dblCHeptane As Double
  Private dblCOctane As Double
  Private dblCNonane As Double
  Private dblCCarbonMonoxide As Double
  Private dblCEthylene As Double
  Private dblCHydrogenSulfide As Double
  Private dblCHydrogen As Double

  Public Property Get ErrorFlags() As GasCompositionErrors
    If isDirty Then
       CalculateProperties
    End If
    ErrorFlags = gceErrors
  End Property

  Public Property Get ErrorNotes() As String
    If isDirty Then
       CalculateProperties
    End If
    ErrorNotes = strErrorNotes
  End Property

  Public Property Let BasePressure(ByVal value As Double)
    If Not dblBasePressue = value Then
      dblBasePressue = value
      isDirty = True
    End If
  End Property
  
  Public Property Get BasePressure() As Double
    BasePressure = dblBasePressue
  End Property

  Public Property Let CompositionName(ByVal value As String)
    strCompositionName = value
  End Property
  
  Public Property Get CompositionName() As String
    CompositionName = strCompositionName
  End Property

  Public Property Let CompositionType(ByVal value As GasCompositionType)
    If Not gctCompositionType = value Then
      isDirty = True
      gctCompositionType = value
    End If
  End Property

  Public Property Let CarbonDioxide(ByVal value As Double)
    If Not value = dblCarbonDioxide Then
      isDirty = True
      dblCarbonDioxide = value
    End If
  End Property
  
  Public Property Get CarbonDioxide() As Double
     CarbonDioxide = dblCarbonDioxide
  End Property

  Public Property Let CarbonMonoxide(ByVal value As Double)
    If Not value = dblCarbonMonoxide Then
      isDirty = True
      dblCarbonMonoxide = value
    End If
  End Property
  
  Public Property Get CarbonMonoxide() As Double
     CarbonMonoxide = dblCarbonMonoxide
  End Property

  Public Property Let Ethane(ByVal value As Double)
    If Not value = dblEthane Then
      isDirty = True
      dblEthane = value
    End If
  End Property
  
  Public Property Get Ethane() As Double
     Ethane = dblEthane
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

  Public Property Let Helium(ByVal value As Double)
    If Not value = dblHelium Then
      isDirty = True
      dblHelium = value
    End If
  End Property
  
  Public Property Get Helium() As Double
     Helium = dblHelium
  End Property

  Public Property Let Heptane(ByVal value As Double)
    If Not value = dblHeptane Then
      isDirty = True
      dblHeptane = value
    End If
  End Property
  
  Public Property Get Heptane() As Double
     Heptane = dblHeptane
  End Property

  Public Property Let Hexane(ByVal value As Double)
    If Not value = dblHexane Then
      isDirty = True
      dblHexane = value
    End If
  End Property
  
  Public Property Get Hexane() As Double
     Hexane = dblHexane
  End Property

  Public Property Let Hydrogen(ByVal value As Double)
    If Not value = dblHydrogen Then
      isDirty = True
      dblHydrogen = value
    End If
  End Property
  
  Public Property Get Hydrogen() As Double
     Hydrogen = dblHydrogen
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

  Public Property Let IsoButane(ByVal value As Double)
    If Not value = dblIsoButane Then
      isDirty = True
      dblIsoButane = value
    End If
  End Property
  
  Public Property Get IsoButane() As Double
     IsoButane = dblIsoButane
  End Property

  Public Property Let IsoPentane(ByVal value As Double)
    If Not value = dblIsoPentane Then
      isDirty = True
      dblIsoPentane = value
    End If
  End Property
  
  Public Property Get IsoPentane() As Double
     IsoPentane = dblIsoPentane
  End Property

  Public Property Let Methane(ByVal value As Double)
    If Not value = dblMethane Then
      isDirty = True
      dblMethane = value
    End If
  End Property
  
  Public Property Get Methane() As Double
     Methane = dblMethane
  End Property

  Public Property Let NeoPentane(ByVal value As Double)
    If Not value = dblNeoPentane Then
      isDirty = True
      dblNeoPentane = value
    End If
  End Property
  
  Public Property Get NeoPentane() As Double
     NeoPentane = dblNeoPentane
  End Property

  Public Property Let Nonane(ByVal value As Double)
    If Not value = dblNonane Then
      isDirty = True
      dblNonane = value
    End If
  End Property
  
  Public Property Get Nonane() As Double
     Nonane = dblNonane
  End Property

  Public Property Let NormalButane(ByVal value As Double)
    If Not value = dblNormalButane Then
      isDirty = True
      dblNormalButane = value
    End If
  End Property
  
  Public Property Get NormalButane() As Double
     NormalButane = dblNormalButane
  End Property

  Public Property Let NormalPentane(ByVal value As Double)
    If Not value = dblNormalPentane Then
      isDirty = True
      dblNormalPentane = value
    End If
  End Property
  
  Public Property Get NormalPentane() As Double
     NormalPentane = dblNormalPentane
  End Property

  Public Property Let Octane(ByVal value As Double)
    If Not value = dblOctane Then
      isDirty = True
      dblOctane = value
    End If
  End Property
  
  Public Property Get Octane() As Double
     Octane = dblOctane
  End Property

  Public Property Let Oxygen(ByVal value As Double)
    If Not value = dblOxygen Then
      isDirty = True
      dblOxygen = value
    End If
  End Property
  
  Public Property Get Oxygen() As Double
     Oxygen = dblOxygen
  End Property

  Public Property Let Propane(ByVal value As Double)
    If Not value = dblPropane Then
      isDirty = True
      dblPropane = value
    End If
  End Property
  
  Public Property Get Propane() As Double
     Propane = dblPropane
  End Property

  Public Property Let Water(ByVal value As Double)
    If Not value = dblWater Then
      isDirty = True
      dblWater = value
    End If
  End Property
  
  Public Property Get Water() As Double
     Water = dblWater
  End Property

  Public Property Get THCToFreeInertRatio() As Double
    If isDirty Then
      CalculateProperties
    End If
    THCToFreeInertRatio = dblTHCToFreeInertRatio
  End Property

  Public Property Get MethaneNumber() As Double
    If isDirty Then
      CalculateProperties
    End If
    MethaneNumber = dblMethaneNumber
  End Property

  Public Property Get SpecificGravity() As Double
    If isDirty Then
      CalculateProperties
    End If
    SpecificGravity = dblSPGrav
  End Property

  Public Property Get HigherHeatingValue() As Double
    If isDirty Then
      CalculateProperties
    End If
    HigherHeatingValue = dblHHV
  End Property

  Public Property Get LowerHeatingValue() As Double
    If isDirty Then
      CalculateProperties
    End If
    LowerHeatingValue = dblLHV
  End Property

  Public Property Get StoichAFRV() As Double
    If isDirty Then
      CalculateProperties
    End If
    StoichAFRV = dblStoichAFR
  End Property

  Public Property Get StoichAFRM() As Double
    If isDirty Then
      CalculateProperties
    End If
    StoichAFRM = dblMSTOICH
  End Property

  Public Property Get WobbeIndex() As Double
    If isDirty Then
      CalculateProperties
    End If
    WobbeIndex = dblWobbeIndex
  End Property

  Public Property Get H2SAdjustment() As Double
    If isDirty Then
      CalculateProperties
    End If
    H2SAdjustment = dblH2SAdj
  End Property

  Public Property Get HeavyHCAdjustment() As Double
    If isDirty Then
      CalculateProperties
    End If
    HeavyHCAdjustment = heavyHCMNAdj
  End Property

  Public Property Get CO2Adjustment() As Double
    If isDirty Then
      CalculateProperties
    End If
    CO2Adjustment = dblCO2Adj
  End Property

  Public Sub CalculateProperties()

    Dim compMult As Double
    Dim stoichFactor As Double
    Dim rawHHV As Double
    Dim rawLHV As Double
    Dim compressibility As Double 'gas compressibility
    
    Dim nonMethaneCombustables As Double
    Dim zfactor As Double
    'clear errors
    gceErrors = noErrors

    Dim rawComponentSum As Double
    rawComponentSum = dblMethane + dblEthane + dblPropane + dblIsoButane + dblNormalButane + dblIsoPentane + dblNormalPentane _
      + dblNeoPentane + dblHexane + dblHeptane + dblOctane + dblHeptane + dblCarbonMonoxide _
      + dblHydrogen + dblHydrogenSulfide + dblOxygen + dblCarbonDioxide + dblNitrogen + dblHelium + dblWater
    
    'normalize the components
    If rawComponentSum <= 0 Then
      gceErrors = compositionZero
    End If
    If Not gceErrors = compositionZero Then
      If gctCompositionType = moleFraction Then
        If rawComponentSum < 0.99999999 Then
          gceErrors = compositionSumLow
        ElseIf rawComponentSum > 1.00000001 Then
          gceErrors = compositionSumHigh
        End If
      Else
        If rawComponentSum < 99.999999 Then
          gceErrors = compositionSumLow
        ElseIf rawComponentSum > 100.000001 Then
          gceErrors = compositionSumHigh
        End If
      End If
      'normalize components
      compMult = 1 / rawComponentSum
      dblC1 = dblMethane * compMult
      dblC2 = dblEthane * compMult
      dblC3 = dblPropane * compMult
      dblIC4 = dblIsoButane * compMult
      dblNC4 = dblNormalButane * compMult
      dblIC5 = dblIsoPentane * compMult
      dblNC5 = dblNormalPentane * compMult
      dblNEC5 = dblNeoPentane * compMult
      dblC6 = dblHexane * compMult
      dblC7 = dblHeptane * compMult
      dblC8 = dblOctane * compMult
      dblC9 = dblNonane * compMult
      dblCO = dblCarbonMonoxide * compMult
      dblCO2 = dblCarbonDioxide * compMult
      dblH2 = dblHydrogen * compMult
      dblO2 = dblOxygen * compMult
      dblN2 = dblNitrogen * compMult
      dblHE = dblHelium * compMult
      dblH2S = dblHydrogenSulfide * compMult
      dblH2O = dblWater * compMult
      
      'calculate heating values
      rawLHV = dblC1 * 909.4 + dblC2 * 1618.7 + dblC3 * 2314.9 + dblIC4 * 3000.4 + dblNC4 * 3010.8 _
        + dblIC5 * 3699 + dblNC5 * 3703.9 + dblNEC5 * 3683 + dblC6 * 4403.9 + dblC7 * 5100.3 _
        + dblC8 * 5796.2 + dblC9 * 6493.6 + dblCO * 320.5 + dblH2 * 273.93 + dblH2S * 586.8
      rawHHV = dblC1 * 1010 + dblC2 * 1769.7 + dblC3 * 2516.1 + dblIC4 * 3251.9 + dblNC4 * 3262.3 _
        + dblIC5 * 4000.9 + dblNC5 * 4008.9 + dblNEC5 * 3985 + dblC6 * 4755.9 + dblC7 * 5502.5 _
        + dblC8 * 6248.9 + dblC9 * 6996.5 + dblCO * 320.5 + dblH2 * 324.2 + dblH2S * 637.1
      zfactor = dblC1 * 0.0116 + dblC2 * 0.0239 + dblC3 * 0.0344 + dblIC4 * 0.0458 + dblNC4 * 0.0478 _
        + dblIC5 * 0.0581 + dblNC5 * 0.0631 + dblNEC5 * 0 + dblC6 * 0.0802 + dblC7 * 0.0944 _
        + dblC8 * 0.1137 + dblC9 * 0.1331 + dblCO * 0.0053 + dblCO2 * 0.0197 + dblH2 * 0 + dblO2 * 0.0073 _
        + dblN2 * 0.0044 + dblH2S * 0.0253 + dblH2O * 0.0623
      compressibility = 1 - zfactor ^ 2 * dblBasePressue
      dblLHV = rawLHV / compressibility
      dblHHV = rawHHV / compressibility
      
      'calculate air/fuel ratios
      dblStoichAFR = dblC1 * 9.528 + dblC2 * 16.675 + dblC3 * 23.821 + (dblIC4 + dblNC4) * 30.967 _
        + (dblIC5 + dblNC5 + dblNEC5) * 38.114 + dblC6 * 45.26 + dblC7 * 52.406 + dblC8 * 59.552 _
        + dblC9 * 66.7 + dblCO * 2.382 + dblH2 * 2.382 + dblH2S * 7.146
      Dim oxygenAdjustment As Double
      Dim oxygenAdjustmentFactor As Double
      oxygenAdjustmentFactor = 1
      If Not (dblO2 > 0 And dblN2 > 0) Then
        stoichFactor = rawLHV / (dblStoichAFR + 1)
      Else
        If dblO2 * 3.764 > dblN2 Then
          oxygenAdjustment = dblN2 * 1.26567481402763
        Else
          oxygenAdjustment = dblO2 * 4.764
        End If
        oxygenAdjustmentFactor = (1 - oxygenAdjustment)
        If oxygenAdjustmentFactor <> 0 Then
          stoichFactor = rawLHV / oxygenAdjustmentFactor / (dblStoichAFR / oxygenAdjustmentFactor + 1)
        Else
          stoichFactor = 0
        End If
      End If
      'calculate specific gravity
      dblSPGrav = dblC1 * 0.55392 + dblC2 * 1.0382 + dblC3 * 1.5226 + dblIC4 * 2.0068 + dblNC4 * 2.0068 _
        + (dblIC5 + dblNC5 + dblNEC5) * 2.4912 + dblC6 * 2.9755 + dblC7 * 3.4598 + dblC8 * 3.9441 + dblC9 * 4.4284 _
        + dblCO * 0.96711 + dblCO2 * 1.5196 + dblH2 * 0.0696 + dblO2 * 1.1048 + dblN2 * 0.96723 + dblHE * 0.1382 _
        + dblH2S * 1.1767 + dblH2O * 0.62202
      'calculate Wobbe index
      dblWobbeIndex = dblHHV / dblSPGrav ^ 0.5
      'calculate mehane adjustment
      Dim nonCombustables As Double
      dblInerts = dblCO2 + dblO2 + dblN2 + dblHE + dblH2O
      Dim allCombustablesFraction As Double
      allCombustablesFraction = (1 - dblInerts)
      
      If allCombustablesFraction = 0 Then
        'all inerts
        gceErrors = gceErrors Or allInerts
        dblMethaneNumber = -9999
      Else
        Dim cC1 As Double
        Dim cC2 As Double
        Dim cC3 As Double
        Dim cC4 As Double
        Dim cC5 As Double
        Dim cC6 As Double
        Dim cC7 As Double
        Dim cC8 As Double
        Dim cC9 As Double
        Dim cCO As Double
        Dim cH2 As Double
        'normalize combustables, lump hydrogen components w/ methane
        cC1 = (dblC1 + dblH2 + dblH2S) / allCombustablesFraction
        cC2 = dblC2 / allCombustablesFraction
        cC3 = dblC3 / allCombustablesFraction
        cC4 = (dblIC4 + dblNC4) / allCombustablesFraction
        cC5 = (dblIC5 + dblNC5 + dblNEC5) / allCombustablesFraction
        cC6 = dblC6 / allCombustablesFraction
        cC7 = dblC7 / allCombustablesFraction
        cC8 = dblC8 / allCombustablesFraction
        cCO = dblC0 / allCombustablesFraction
        cH2 = dblH2 / allCombustablesFraction
        Dim carbonSum As Double
        Dim hydrogenSum As Double
        carbonSum = cC1 + cC2 * 2 + cC3 * 3 + cC4 * 4 + cC5 * 5 + cC6 * 6 + cC7 * 7 + cC8 * 8 + cC9 * 9 + cCO
        hydrogenSum = cC1 * 4 + cC2 * 6 + cC3 * 8 + cC4 * 10 + cC5 * 12 + cC6 * 14 + cC7 * 16 + cC8 * 18 + cC9 * 20
        If carbonSum = 0 Then
          gceErrors = gceErrors Or allHydrogen
          mnAdj = 0
        ElseIf hydrogenSum = 0 Then
          gceErrors = gceErrors Or allCarbon
          mnAdj = -9999
        Else
          Dim hcrAdj As Double
          Dim sgAdj As Double
          hcrAdj = hydrogenSum / carbonSum
          sgAdj = cC1 * 0.55392 + cC2 * 1.0382 + cC3 * 1.5226 + cC4 * 2.0068 + cC5 * 2.4912 + cC6 * 2.9755 _
            + cC7 * 3.4598 + cC8 * 3.9441 + cC9 * 4.4284 + cCO * 0.96711
          dblMNRaw = 34.26085951 * hcrAdj + 2.944058733E-09 * sgAdj * (21.30716314 * hcrAdj ^ 2) ^ hcrAdj - 59.1854841
        End If
          
        CalcCO2Adjustment
        CalcH2SAdjustment
        dblMethaneNumber = dblMNRaw + dblCO2Adj + dblH2SAdj - cH2 * 100
      End If 'all inerts
    End If 'compositionzero
    'build error string
    If gceErrors = noErrors Then
      strErrorNotes = "No calculation errors or warnings."
    Else
      strErrorNotes = ""
      If (gceErrors And GasCompositionErrors.compositionZero) > 0 Then
        strErrorNotes = "Gas compositions are less than or equal to zero."
      End If
      If (gceErrors And GasCompositionErrors.compositionSumHigh) > 0 Then
        If gctCompositionType = moleFraction Then
          strErrorNotes = "Gas compositions sum to more than 1.0; "
        Else
          strErrorNotes = "Gas compositions sum to more than 100.0; "
        End If
      End If
      If (gceErrors And GasCompositionErrors.compositionSumLow) > 0 Then
        If gctCompositionType = moleFraction Then
          strErrorNotes = "Gas compositions sum to less than 1.0; "
        Else
          strErrorNotes = "Gas compositions sum to less than 100.0; "
        End If
      End If
      If (gceErrors And GasCompositionErrors.allCarbon) > 0 Then
        strErrorNotes = strErrorNotes & "Gas contains all carbon components; "
      End If
      If (gceErrors And GasCompositionErrors.allCarbon) > 0 Then
        strErrorNotes = strErrorNotes & "Gas contains all carbon components; "
      End If
      If (gceErrors And GasCompositionErrors.allCarbon) > 0 Then
        strErrorNotes = strErrorNotes & "Gas contains all hydrogen components; "
      End If
      If (gceErrors And GasCompositionErrors.allInerts) > 0 Then
        strErrorNotes = strErrorNotes & "Gas contains all inerts; "
      End If
    End If
    
    isDirty = False
  End Sub

  Private Sub CalcH2SAdjustment()
    If dblCHydrogenSulfide <> 0 Then
      dblH2SAdj = -14.83 * dblH2S - 1.048 * dblH2S ^ 0.002036
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
      dblCO2Adj = 57.0952 * dblCO2 ^ 2 * mnAdj + 0.0000011611 * dblCO2 * mnAdj ^ 4 - 0.081104 - 45.2077 * dblCO2 - 2153.3903 * dblCO2 ^ 2 - 253.9675 * dblCO2 ^ 3 - 0.337730332431225 * dblCO2 ^ 2 * mnAdj ^ 2
    End If
  End Sub

Private Sub Class_Initialize()
  dblBasePressue = 14.73
End Sub
