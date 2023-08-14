# dependencies
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.Wpf')
# find the path of ui.xaml
from pyrevit import UI
from pyrevit import script

xamlfile = script.get_bundle_file('ui.xaml')
# import WPF creator and base Window
import wpf
from System import Windows
# revit api
from pyrevit import revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
import math


uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = revit.get_selection()


class CustomISelectionFilter(UI.Selection.ISelectionFilter):
    # standard API override function
    def __init__(self, category_name):
        self.category_name = category_name

    def AllowElement(self, e):
        if e.Category.Name == self.category_name:
            return True
        else:
            return False

    # standard API override function
    def AllowReference(self, refer, point):
        return False


def UnitConversion(value, get_internal=False, units="m"):
    """ Converts to or from internal units """
    from Autodesk.Revit.DB import UnitTypeId
    from Autodesk.Revit.DB import UnitUtils
    if units == "m":
        units = UnitTypeId.Meters
    elif units == "m2":
        units = UnitTypeId.SquareMeters
    elif units == "cm":
        units = UnitTypeId.Centimeters
    elif units == "ft":
        units = UnitTypeId.Feet
    elif units == "mm":
        units = UnitTypeId.Millimeters
    if get_internal:
        return UnitUtils.ConvertToInternalUnits(value, units)
    return UnitUtils.ConvertFromInternalUnits(value, units)


def CurveDivisions(crvs, div):
    """ This function returns dictionary of curves and its points divided by the given input """
    mydict = {}
    for curve in crvs:
        mydict[curve] = []
        for i in range(div):
            x = i + 1
            point = curve.Evaluate((x / div), True)
            mydict[curve].append(point)
        point = curve.Evaluate((0), True)
        mydict[curve].append(point)
    return mydict

def CheckForDupCoord(listOfCoords):
    """ Checks the list for duplicate coordinates and return a new list without duplicates """
    listOfCoordsWithoutDup = []
    for pnt in listOfCoords:
        if len(listOfCoordsWithoutDup) == 0:
            listOfCoordsWithoutDup.append(pnt)
        else:
            for p in listOfCoordsWithoutDup:
                checkForDup = False
                if p.IsAlmostEqualTo(pnt):
                    checkForDup = True
                    break
            if checkForDup == False:
                listOfCoordsWithoutDup.append(pnt)
    return listOfCoordsWithoutDup


class MyFailureProcessor(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        return FailureProcessingResult.Continue


class MyWindows(Windows.Window):
    def __init__(self):
        wpf.LoadComponent(self, xamlfile)

    def btnCreate_Click(self, sender, args):
        """ Get inputs from user """
        self.Close()
        angle = int(self.angleParam.Text)
        offSet = int(self.offsetParam.Text)
        segment_length = int(self.segment_length.Text)

        MyWindows().runScript(angle, offSet, segment_length)

    def get_3D_view(self):
        """ Function to get ViewType - 3D View """
        views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)  # Get 3D view
        for view in views:
            if view.Name == "{3D}":
                return view

    def offset_curves_from_BP(self, buildingPad, offsetParam):
        """ Offsets boundary cloops of BP """
        curves = []  # Offset curves from foundation pad
        try:  # Get curves from foundation pad, and create model lines
            opt = Options()
            opt.View = self.get_3D_view()
            elemGeo = buildingPad.get_Geometry(opt)
            # INPUT

            faceBP = None
            for geo in elemGeo:  # Get surface with normal Z == 1, and offset its curves
                faces = geo.Faces
                faceBP = None
                for face in faces:
                    if face.FaceNormal.Z == 1:
                        faceBP = face
                        pass

            cLoops = faceBP.GetEdgesAsCurveLoops()[0]
            offsetParam_to_FT = UnitConversion(offsetParam, True, "mm")
            offSetCLoops = CurveLoop.CreateViaOffset(cLoops, offsetParam_to_FT, XYZ(0, 0, 1))
            t = Transaction(doc, "Model Lsssines")  # Create model lines form curves for visual purposes
            t.Start()
            for curve in offSetCLoops:  # Create model lines from offset curves
                #sPlane = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, curve.Origin))
                #doc.Create.NewModelCurve(curve, sPlane)
                curves.append(curve)
            t.Commit()

        except Exception as e:
            print(e)

        return curves, offSetCLoops

    def CurveDivisionSingular(self, crvs, div):
        """ This function returns a list of points divided by divisions """
        points = []
        point0 = crvs.Evaluate(0, True)  # Starting point of a curve
        points.append(point0)
        for i in range(div):
            xEnum = i + 1.0
            point = crvs.Evaluate((xEnum / div), True)
            points.append(point)
        return points


    def create_rotated_lines(self, curves, offSettedCLoops, angleExcavation, segment_length):
        """ Creates rotated curves from user input """
        t = Transaction(doc, "Create model lines")
        t.Start()

        # Input in degrees for rotation
        rotationInRadians = (90 - angleExcavation) * math.pi / 180

        distance = None
        zCoor_top = None
        zCoor_bot = None

        for c in curves:
            point = c.Evaluate(0.5, True)

            xCo = point.X
            yCo = point.Y
            zCo = point.Z
            newPoint = XYZ(xCo, yCo, 10)  # New vertical line with arbitrary length
            newLine = Line.CreateBound(point, newPoint)  # Create a line from a point of the offset curve with dir. Z+

            # Create rotated line
            rotation = Transform.CreateRotationAtPoint(c.Direction, rotationInRadians, c.Origin)
            rotatedLine = newLine.CreateTransformed(rotation)

            #  Get end points of rotated lines
            zCoor_top = rotatedLine.GetEndPoint(1).Z
            zCoor_bot = zCo
            distance1 = rotatedLine.GetEndPoint(1).DistanceTo(XYZ(xCo, yCo, zCoor_top))
            distance = distance1

            if distance or zCoor_top or zCoor_bot != None:
                break

        t.Commit()
        offSetCLoopsTop = CurveLoop.CreateViaOffset(offSettedCLoops, distance, XYZ(0, 0, 1))  # at the level of foundation pad
        movedOffsetCLoopsTop = []
        # Create model lines form curves for visual purposes
        t = Transaction(doc, "Model Lines")
        t.Start()

        for curve in offSetCLoopsTop:
            dis = XYZ(0, 0, zCoor_top).DistanceTo(XYZ(0, 0, zCoor_bot))
            translation = Transform.CreateTranslation(XYZ(0, 0, dis))
            translated_line = curve.CreateTransformed(translation)
            movedOffsetCLoopsTop.append(translated_line)
            # Create model lines from offset curves
            #sPlane = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, translated_line.Origin))
            #doc.Create.NewModelCurve(translated_line, sPlane)

        t.Commit()

        t = Transaction(doc, "Create model lines")
        t.Start()
        #  This part creates sloped curves
        slopeCurves = []
        for c1, c2 in zip(offSettedCLoops, movedOffsetCLoopsTop):  # They do correlate
            length = UnitUtils.ConvertFromInternalUnits(c1.Length, UnitTypeId.Millimeters)
            segments = int(length // segment_length)
            points1 = self.CurveDivisionSingular(c1, segments)
            points2 = self.CurveDivisionSingular(c2, segments)
            point3 = c1.Evaluate(0.521211, True)  # Arbitrary point on the line to create a sketch plane

            for p1, p2 in zip(points1, points2):
                #  Create a new line
                newLine = Line.CreateBound(p1, p2)
                #sPlane = SketchPlane.Create(doc, Plane.CreateByThreePoints(p1, p2, point3))
                #doc.Create.NewModelCurve(newLine, sPlane)
                slopeCurves.append(newLine)
        t.Commit()
        return slopeCurves

    def ProjectPointsOnTopographySurface(self, curves, topography, segment_length):
        """ This function returns intersection points from lines in a topography, and a list of segment points"""
        referenceIntersector = ReferenceIntersector(topography.Id, FindReferenceTarget.Mesh, self.get_3D_view())
        points_list = []
        segment_points_list = []
        curve_from_basepoint_to_intersection = []
        for c in curves:
            curveDir = c.Direction
            curveOrig = c.Origin
            # intersectionFilter = Autodesk.Revit.DB.ElementClassFilter(typeTopography)
            # referenceIntersector = ReferenceIntersector(intersectionFilter, FindReferenceTarget.All, view3D)
            referenceWithContext = referenceIntersector.FindNearest(curveOrig, curveDir)
            intersection_points = referenceWithContext.GetReference().GlobalPoint
            points_list.append(intersection_points)
            new_line = Line.CreateBound(c.GetEndPoint(0), intersection_points)
            curve_from_basepoint_to_intersection.append(new_line)

        for c in curve_from_basepoint_to_intersection:
            length = UnitUtils.ConvertFromInternalUnits(c.Length, UnitTypeId.Millimeters)
            segments = int(length // segment_length)  # Create point every 200 mm
            segment_points = self.CurveDivisionSingular(c, segments)
            for p in segment_points:
                segment_points_list.append(p)

        return points_list, segment_points_list

    def check_intersecting_points(self, curve_list, topography, segment_length):
        """ Returns intersection points from provided curve list """
        prelimList = []
        # Get foundation base offset points
        for c in curve_list:
            prelimList.append(c.GetEndPoint(0))
        # Get intersection points from sloped Curves
        intersectionPoints, segment_points = self.ProjectPointsOnTopographySurface(curve_list, topography, segment_length)
        intersectionPoints_NoDups = CheckForDupCoord(intersectionPoints + prelimList + segment_points)
        # List of all new points to be added to the existing topography
        return intersectionPoints_NoDups

    def update_points_on_topography(self, points, topography):
        try:
            ts = Architecture.TopographyEditScope(doc, "Edit topo points")
            ts.Start(topography.Id)
            t = Transaction(doc, "Edit topo points")
            t.Start()
            topography.AddPoints(points)
            t.Commit()
            ts.Commit(MyFailureProcessor())
        except Exception as e:
            print(e)

    def runScript(self, angleExcavation=45, offsetExcavation=800, segment_length=200):
        sel_Topo = revit.uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Topography"),
                                                    "Select a Topography")
        topography = doc.GetElement(sel_Topo)
        sel_BP = revit.uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter("Pads"),
                                                  "Select a Building pad")
        buildingPad = doc.GetElement(sel_BP)

        boundaryCurves, offSetCLoops = self.offset_curves_from_BP(buildingPad, offsetExcavation)

        rotated_curves_list = self.create_rotated_lines(boundaryCurves, offSetCLoops, angleExcavation, segment_length)
        points_for_topography = self.check_intersecting_points(rotated_curves_list, topography, segment_length)
        self.update_points_on_topography(points_for_topography, topography)

# Show the window
if __name__ == '__main__':
    GUI = MyWindows().ShowDialog()
