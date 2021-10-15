VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Générateur de courbes
' TP Info, Lancelot Herrbach 76Ch220
' 2021 ENSAM GIE

Dim coords(3, 2) As Double

Sub NewPart()

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Add("Part")

End Sub

Sub PlacePoints()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Add()

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim hybridShapePointCoord1 As HybridShapePointCoord
Dim axisSystems1 As AxisSystems
Dim axisSystem1 As AxisSystem
Dim reference1 As Reference

For i = 0 To 3
        
        Debug.Print (coords(i, 0))
        
        Set hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(coords(i, 0), coords(i, 1), coords(i, 2))
        
        Set axisSystems1 = part1.AxisSystems
        Set axisSystem1 = axisSystems1.Item("Repère absolu")
        Set reference1 = part1.CreateReferenceFromObject(axisSystem1)
        hybridShapePointCoord1.RefAxisSystem = reference1
    
        hybridBody1.AppendHybridShape hybridShapePointCoord1
        
        part1.InWorkObject = hybridShapePointCoord1
        
        part1.Update
    

Next

End Sub

Sub TraceSpline()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim hybridShapeSpline1 As HybridShapeSpline
Set hybridShapeSpline1 = hybridShapeFactory1.AddNewSpline()

hybridShapeSpline1.SetSplineType 0

hybridShapeSpline1.SetClosing 0

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Set géométrique.1")

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim hybridShapePointCoord1 As HybridShapePointCoord
Set hybridShapePointCoord1 = hybridShapes1.Item("Point.1")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)

hybridShapeSpline1.AddPointWithConstraintExplicit reference1, Nothing, -1#, 1, Nothing, 0#

Dim hybridShapePointCoord2 As HybridShapePointCoord
Set hybridShapePointCoord2 = hybridShapes1.Item("Point.2")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(hybridShapePointCoord2)

hybridShapeSpline1.AddPointWithConstraintExplicit reference2, Nothing, -1#, 1, Nothing, 0#

Dim hybridShapePointCoord3 As HybridShapePointCoord
Set hybridShapePointCoord3 = hybridShapes1.Item("Point.3")

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord3)

hybridShapeSpline1.AddPointWithConstraintExplicit reference3, Nothing, -1#, 1, Nothing, 0#

Dim hybridShapePointCoord4 As HybridShapePointCoord
Set hybridShapePointCoord4 = hybridShapes1.Item("Point.4")

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(hybridShapePointCoord4)

hybridShapeSpline1.AddPointWithConstraintExplicit reference4, Nothing, -1#, 1, Nothing, 0#

hybridBody1.AppendHybridShape hybridShapeSpline1

part1.InWorkObject = hybridShapeSpline1

part1.Update

End Sub


Sub AddRefPlane()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Set géométrique.1")

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim hybridShapeSpline1 As HybridShapeSpline
Set hybridShapeSpline1 = hybridShapes1.Item("Spline.1")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(hybridShapeSpline1)

Dim hybridShapePointCoord1 As HybridShapePointCoord
Set hybridShapePointCoord1 = hybridShapes1.Item("Point.1")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(hybridShapePointCoord1)

Dim hybridShapePlaneNormal1 As HybridShapePlaneNormal
Set hybridShapePlaneNormal1 = hybridShapeFactory1.AddNewPlaneNormal(reference1, reference2)

hybridBody1.AppendHybridShape hybridShapePlaneNormal1

part1.InWorkObject = hybridShapePlaneNormal1

part1.Update

End Sub

Sub TraceSketchCercle()


Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Set géométrique.1")

Dim sketches1 As Sketches
Set sketches1 = hybridBody1.HybridSketches

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim reference1 As Reference
Set reference1 = hybridShapes1.Item("Plan.1")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 21.998008
arrayOfVariantOfDouble1(1) = 20.447643
arrayOfVariantOfDouble1(2) = -4.035123
arrayOfVariantOfDouble1(3) = -0.680825
arrayOfVariantOfDouble1(4) = 0.732446
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0.09753
arrayOfVariantOfDouble1(7) = 0.090657
arrayOfVariantOfDouble1(8) = 0.991095
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 2

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(15.268373, 4.071379)

point2D1.ReportName = 3

Dim circle2D1 As Circle2D
Set circle2D1 = factory2D1.CreateClosedCircle(15.268373, 4.071379, InParam1.Value)

circle2D1.CenterPoint = point2D1

circle2D1.ReportName = 4

Dim reference2 As Reference
Set reference2 = hybridShapes1.Item("Point.1")

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference2)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Empreinte.1")

geometry2D1.Construction = True

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(point2D1)

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(geometry2D1)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeOn, reference3, reference4)

constraint1.Mode = catCstModeDrivingDimension

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddMonoEltCst(catCstTypeRadius, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint2.Dimension

length1.Value = InParam1.Value

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.Update


End Sub

Sub TraceSketechElipse()


Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Set géométrique.1")

Dim sketches1 As Sketches
Set sketches1 = hybridBody1.HybridSketches

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim reference1 As Reference
Set reference1 = hybridShapes1.Item("Plan.1")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 0#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = -0.984966
arrayOfVariantOfDouble1(4) = 0.172747
arrayOfVariantOfDouble1(5) = -0#
arrayOfVariantOfDouble1(6) = -0.150953
arrayOfVariantOfDouble1(7) = -0.860697
arrayOfVariantOfDouble1(8) = 0.486224
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 2

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(0#, 0#)

point2D1.ReportName = 3

Dim ellipse2D1 As Ellipse2D
Set ellipse2D1 = factory2D1.CreateClosedEllipse(0#, 0#, 37.651546, -5.134302, 38#, 12#)

ellipse2D1.CenterPoint = point2D1

ellipse2D1.ReportName = 4

Dim reference2 As Reference
Set reference2 = hybridShapes1.Item("Point.1")

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference2)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Empreinte.1")

geometry2D1.Construction = True

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(point2D1)

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(geometry2D1)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeOn, reference3, reference4)

constraint1.Mode = catCstModeDrivingDimension

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(ellipse2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddMonoEltCst(catCstTypeMajorRadius, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint2.Dimension

length1.Value = InParam1.Value

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(ellipse2D1)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddMonoEltCst(catCstTypeMinorRadius, reference6)

constraint3.Mode = catCstModeDrivingDimension

Dim length2 As Length
Set length2 = constraint3.Dimension

length2.Value = InParam2.Value

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.Update


End Sub


Sub TraceSketchRect()


Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Set géométrique.1")

Dim sketches1 As Sketches
Set sketches1 = hybridBody1.HybridSketches

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim reference1 As Reference
Set reference1 = hybridShapes1.Item("Plan.1")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 0#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = -0.984966
arrayOfVariantOfDouble1(4) = 0.172747
arrayOfVariantOfDouble1(5) = -0#
arrayOfVariantOfDouble1(6) = -0.150953
arrayOfVariantOfDouble1(7) = -0.860697
arrayOfVariantOfDouble1(8) = 0.486224
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 2

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(0#, 0#)

point2D1.ReportName = 3

Dim point2D2 As Point2D
Set point2D2 = factory2D1.CreatePoint(38#, 12#)

point2D2.ReportName = 4

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(38#, -12#)

point2D3.ReportName = 5

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(38#, 12#, 38#, -12#)

line2D3.ReportName = 6

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(-38#, -12#)

point2D4.ReportName = 7

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(38#, -12#, -38#, -12#)

line2D4.ReportName = 8

line2D4.StartPoint = point2D3

line2D4.EndPoint = point2D4

Dim point2D5 As Point2D
Set point2D5 = factory2D1.CreatePoint(-38#, 12#)

point2D5.ReportName = 9

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(-38#, -12#, -38#, 12#)

line2D5.ReportName = 10

line2D5.StartPoint = point2D4

line2D5.EndPoint = point2D5

Dim line2D6 As Line2D
Set line2D6 = factory2D1.CreateLine(-38#, 12#, 38#, 12#)

line2D6.ReportName = 11

line2D6.StartPoint = point2D5

line2D6.EndPoint = point2D2

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(line2D3)

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(line2D2)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeVerticality, reference2, reference3)

constraint1.Mode = catCstModeDrivingDimension

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(line2D4)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(line2D5)

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(line2D2)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeVerticality, reference6, reference7)

constraint3.Mode = catCstModeDrivingDimension

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(line2D6)

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D1)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference8, reference9)

constraint4.Mode = catCstModeDrivingDimension

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(line2D3)

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D5)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(point2D1)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference10, reference11, reference12)

constraint5.Mode = catCstModeDrivingDimension

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(line2D4)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromObject(line2D6)

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(point2D1)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference13, reference14, reference15)

constraint6.Mode = catCstModeDrivingDimension

Dim reference16 As Reference
Set reference16 = hybridShapes1.Item("Point.1")

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference16)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Empreinte.1")

geometry2D1.Construction = True

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromObject(point2D1)

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(geometry2D1)

Dim constraint7 As Constraint
Set constraint7 = constraints1.AddBiEltCst(catCstTypeOn, reference17, reference18)

constraint7.Mode = catCstModeDrivingDimension

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromObject(line2D3)

Dim constraint8 As Constraint
Set constraint8 = constraints1.AddMonoEltCst(catCstTypeLength, reference19)

constraint8.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint8.Dimension

length1.Value = InParam1.Value

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(line2D6)

Dim constraint9 As Constraint
Set constraint9 = constraints1.AddMonoEltCst(catCstTypeLength, reference20)

constraint9.Mode = catCstModeDrivingDimension

Dim length2 As Length
Set length2 = constraint9.Dimension

length2.Value = InParam2.Value

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.Update

End Sub

Sub Extrude()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

part1.InWorkObject = body1

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Set géométrique.1")

Dim sketches1 As Sketches
Set sketches1 = hybridBody1.HybridSketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.1")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(sketch1)

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim hybridShapeSpline1 As HybridShapeSpline
Set hybridShapeSpline1 = hybridShapes1.Item("Spline.1")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(hybridShapeSpline1)

Dim rib1 As Rib
Set rib1 = shapeFactory1.AddNewRibFromRef(reference1, reference2)

part1.Update

End Sub

Sub GetPointsCoords()

coords(0, 0) = InX1.Value
coords(0, 1) = InY1.Value
coords(0, 2) = InZ1.Value

coords(1, 0) = InX2.Value
coords(1, 1) = InY2.Value
coords(1, 2) = InZ2.Value

coords(2, 0) = InX3.Value
coords(2, 1) = InY3.Value
coords(2, 2) = InZ3.Value

coords(3, 0) = InX4.Value
coords(3, 1) = InY4.Value
coords(3, 2) = InZ4.Value

End Sub

Sub UpdateInputs()

Select Case (ComboBox1.ListIndex)
    Case 0
        LabelParam1.Visible = True
        InParam1.Visible = True
        LabelParam1.Caption = "Rayon"
        
        LabelParam2.Visible = False
        LabelParam3.Visible = False
        InParam2.Visible = False
        InParam3.Visible = False
    Case 1
        LabelParam1.Visible = True
        InParam1.Visible = True
        LabelParam1.Caption = "a"
        
        LabelParam2.Visible = True
        InParam2.Visible = True
        LabelParam2.Caption = "b"

        LabelParam3.Visible = False
        InParam3.Visible = False
    Case 2
        LabelParam1.Visible = True
        InParam1.Visible = True
        LabelParam1.Caption = "a"

        LabelParam2.Visible = False
        LabelParam3.Visible = False
        InParam2.Visible = False
        InParam3.Visible = False
    Case 3
        LabelParam1.Visible = True
        InParam1.Visible = True
        LabelParam1.Caption = "a"
        
        LabelParam2.Visible = True
        InParam2.Visible = True
        LabelParam2.Caption = "b"

        LabelParam3.Visible = False
        InParam3.Visible = False
End Select


End Sub

Private Sub ComboBox1_Change()
    UpdateInputs
End Sub

Private Sub CommandButton1_Click()
GetPointsCoords
NewPart
PlacePoints
TraceSpline
AddRefPlane

Select Case (ComboBox1.ListIndex)
    Case 0
        TraceSketchCercle
    Case 1
        TraceSketechElipse
    Case 2
        InParam2.Value = InParam1.Value
        TraceSketchRect
    Case 3
        TraceSketchRect
End Select


Extrude

End Sub

Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Cercle"
    ComboBox1.AddItem "Ellipse"
    ComboBox1.AddItem "Carré"
    ComboBox1.AddItem "Rectangle"
    ComboBox1.ListIndex = 0
    
    UpdateInputs
End Sub
