#! python3
# Import time
import time
start_time = time.time()
from datetime import date, datetime, timedelta
# Import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
# Import
from System.Collections.Generic import List
import os
import json
import requests
# Import Lib
from Snippets._project_path import get_project_size_mb
from Snippets._convert import convert_internal_units
from Snippets.get_journal_path import get_journal_path
from Snippets._get_project_name import get_file_name

doc = __revit__.ActiveUIDocument.Document
today = date.today()
now = datetime.now()

url = "https://api.powerbi.com/beta/d5806acf-9211-41f2-8486-e356d60362b4/datasets/a39ed191-51bc-4b44-aaf0-5dc429ab7eab/rows?cmpid=pbi-glob-head-snn-signin&key=zJ8a86JOMy78%2Fbfl3zep1JAKOJrLTjh5Jdj%2FTSHKWwa40agBtxP81zZpBKHb%2F8RCLdnin5bZ%2FnCVSaIr%2B8SBEA%3D%3D"

def get_purgeable_elements(doc):
    # A constant
    PURGE_GUID = "e8c63650-70b7-435a-9010-ec97660c1bda"
    # A generic list of PerformanceAdviserRuleIds as required by the ExecuteRules method
    rule_id_list = List[PerformanceAdviserRuleId]()
    # Iterating through all PerformanceAdviser rules looking to find that which matches PURGE_GUID
    for rule_id in PerformanceAdviser.GetPerformanceAdviser().GetAllRuleIds():
        if str(rule_id.Guid) == PURGE_GUID:
            rule_id_list.Add(rule_id)
            break
    # Get all purgeable elements
    failure_messages = PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(doc, rule_id_list)
    if len(failure_messages) > 0:
        purgeable_element_ids = failure_messages[0].GetFailingElements()
        return purgeable_element_ids

class RevitHealthCheck():
    """ Revit health check for Power BI dashboard"""
    def __init__(self):
        self.data_parser()

    def general_data(self):
        """ Get general information of the current project """
        user = doc.Application.Username
        software = doc.Application.VersionName
        project = get_file_name(doc)
        size = get_project_size_mb(doc)
        date = now.strftime("%d/%m/%Y %H:%M:%S")
        dateTime = today.strftime("%d/%m/%Y %H:%M:%S")
        dic = {
            "user": user,
            "software": software,
            "project": project,
            "size": size,
            "date": date,
            "dateTime": dateTime
            }

        return dic

    def view_and_sheet_data(self):
        """ Get views and sheets count """
        all_views = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_Views)\
            .WhereElementIsNotElementType()\
            .ToElements()
        all_sheets = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Sheets) \
            .WhereElementIsNotElementType() \
            .ToElements()

        on_sheet = 0                        # Number of placed views
        not_on_sheet = 0                    # Number of unplaced views
        total_views = len(all_views)        # Total view count
        total_sheets = len(all_sheets)      # Number of sheets in the project

        for view in all_views:
            sheet_number = view.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER).AsString()
            if sheet_number == None:
                not_on_sheet += 1
            else:
                on_sheet += 1
        dic = {
            "All views": total_views,
            "All sheets": total_sheets,
            "Views on sheets": on_sheet,
            "Views not on sheets": not_on_sheet
            }

        return dic

    def style_data(self):
        """ Get data about materials, lines and fills"""
        all_materials = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_Materials) \
            .WhereElementIsNotElementType() \
            .ToElements() \

        all_fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
        all_line_patters = FilteredElementCollector(doc).OfClass(LinePatternElement).ToElements()
        all_line_styles = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines).SubCategories.Size

        dic = {
            "Materials": len(all_materials),
            "Line styles": all_line_styles,
            "Line patterns": len(all_line_patters),
            "Fill patterns": len(all_fill_patterns)
        }
        return dic

    def link_import_data(self):
        all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
        all_design_opt = FilteredElementCollector(doc).OfClass(DesignOption).ToElementIds()
        all_imports = FilteredElementCollector(doc).OfClass(ImportInstance).ToElements()
        all_cad_imports = []
        all_cad_links = []
        all_images = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RasterImages).ToElementIds()

        for i in all_imports:
            if i.IsLinked:
                all_cad_links.append(i)
            else:
                all_cad_imports.append(i)

        # Linked Revit files
        all_revit_link_collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        all_revit_links = 0
        all_pinned_rli = 0

        for i in all_revit_link_collector:
            all_revit_links += 1
            if i.Pinned:
                all_pinned_rli += 1

        dic = {
            "Worksets": len(all_worksets),
            "Design options": len(all_design_opt),
            "Imports": len(all_imports),
            "CAD imports": len(all_cad_imports),
            "CAD links": len(all_cad_links),
            "Images": len(all_images),
            "Linked Revit": all_revit_links,
            "Pinned linked Revit": all_pinned_rli
        }

        return dic


    def room_data(self):
        all_rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
        all_rooms_count = len(all_rooms)
        # Room area
        all_rooms_area = 0  # Feet squared

        all_unplaced_rooms = []
        all_not_enclosed_rooms = []

        for r in all_rooms:
            if r.Area == 0:
                if r.Location:
                    all_not_enclosed_rooms.append(r)
                else:
                    all_unplaced_rooms.append(r)
            else:
                all_rooms_area += r.Area

        room_area_meters = convert_internal_units(all_rooms_area, False, "m2")
        dic = {
            "Room count": all_rooms_count,
            "Room area": room_area_meters,
            "Unplaced rooms": len(all_unplaced_rooms),
            "Unenclosed rooms": len(all_not_enclosed_rooms)
        }

        return dic

    def group_data(self):
        """ Get data of groups in the project """

        all_model_groups = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_IOSModelGroups).ToElements()
        all_instances = []
        all_types = []
        all_unused_group_instances = 0

        for i in all_model_groups:
            if str(type(i)) == GroupType:
                all_types.append(i)
            else:
                all_instances.append(i)

        # Get purgable group instances
        f = get_purgeable_elements(doc)
        if f == None:
            pass
        else:
            for i in f:
                ele = doc.GetElement(i)
                if type(ele) == GroupType:
                    all_unused_group_instances += 1
        dic = {
            "Model groups": len(all_model_groups),
            "Model group types": len(all_types),
            "Model group instances": len(all_instances),
            "Unused groups": all_unused_group_instances
        }
        return dic

    def family_data(self):
        """ Get data of families in the project """
        company_abbreviation = "AFRY"
        all_families = FilteredElementCollector(doc).OfClass(Family).ToElements()  # custom families only (not system families)
        all_inplace_families = []  # number of modeled in place families
        company_families = []  # number of families starting with AFRY
        non_company_families = []
        all_unplaced_families = 0  # number of purgeable families

        # Inplace families
        for i in all_families:
            if i.IsInPlace:
                all_inplace_families.append(i)

        # Company families
        for fam in all_families:
            if str(fam.Name).startswith(company_abbreviation):
                company_families.append(fam)
            else:
                non_company_families.append(fam)

        # Warnings
        Warnings = len(doc.GetWarnings())

        # Get purgable family types
        f = get_purgeable_elements(doc)
        if f == None:
            pass
        else:
            for i in f:
                ele = doc.GetElement(i)
                if type(ele) == GroupType:
                    all_unplaced_families += 1
        # Sync time
        sync_dur_int = 0
        if get_journal_path(doc) == None:
            pass
        else:
            sync_dur_int = int(get_journal_path(doc).total_seconds())


        dic = {
            "Loaded families": len(all_families),
            "Inplace families": len(all_inplace_families),
            "Non-approved families": len(non_company_families),
            "Unused families": all_unplaced_families,
            "Warnings": Warnings,
            "Average sync time": sync_dur_int
            }

        return dic

    def data_parser(self):
        """ Parses collected data into JSON """
        general_data = self.general_data()

        general_data.update(self.view_and_sheet_data())
        general_data.update(self.style_data())
        general_data.update(self.link_import_data())
        general_data.update(self.room_data())
        general_data.update(self.group_data())
        general_data.update(self.family_data())

        r = requests.post(url, json=[ general_data ])
        if r.status_code == 200:
            pass
        else:
            txt = "Error occurred: {status_c}, {status_r},".format(status_c=str(r.status_code), status_r=r.reason)
            print(txt)

# Initialize the script
if __name__ == '__main__':
    RevitHealthCheck()
    seconds = time.time() - start_time
    print(str(round(seconds, 2)) + " ms")

