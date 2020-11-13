import fastkml
from zipfile import ZipFile
import tkinter as tk
from tkinter import filedialog
import untangle
import xmltodict
import lxml
from pprint import pprint
import openpyxl


root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
excel_path = filedialog.askopenfilename()

kmz = ZipFile(file_path, 'r')
kml = kmz.open('doc.kml', 'r').read().decode("utf-8")
parsed = untangle.parse(kml)
as_dict = xmltodict.parse(kml)
#print(parsed.kml.Folder.name.cdata)

workbook = openpyxl.load_workbook(filename=excel_path)
sheet = workbook.active

potential_folder = [t for t in as_dict["kml"]["Folder"]["Document"]["Folder"] if t['name']=="Potential New Locations"][0]
potential_sites = potential_folder["Folder"]["Placemark"]

sitenames = []
ids = []
for row in \
    sheet.iter_rows(min_row=2, values_only=True):
    if not row[0] is None:
        ids.append(row[0])
    if not row[2] is None:
        sitenames.append(row[2])

pprint(potential_sites[17])
print(len(potential_sites[17]["ExtendedData"]["SchemaData"]["SimpleData"][1]))

def checkAddress(gis_dict, term):
    if term in list(map(lambda x: x["@name"], gis_site["ExtendedData"]["SchemaData"]["SimpleData"])) \
            and len([t for t in gis_dict["ExtendedData"]["SchemaData"]["SimpleData"] if t["@name"] == term][0]) > 1:
        return [t for t in gis_dict["ExtendedData"]["SchemaData"]["SimpleData"] if t["@name"] == term][0]["#text"]
    else:
        return ""

i = 0
for gis_site in potential_sites:
    i += 1
    print(i)
    if gis_site["name"] not in sitenames:
        # sheet.append(
        #     [max(ids) + 1,
        #      "",
        #      gis_site["name"],
        #      [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if t["@name"] == "Market_Value__Total_"]
        #      ]
        # )
        if "ExtendedData" in gis_site.keys():
            sheet.append(
                [max(ids) + 1,
                 "",
                 gis_site["name"],
                 "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if t["@name"] == "Market_Value__Total_"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if t["@name"] == "Market_Value__Total_"][0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 checkAddress(gis_site, "Address"),
                 checkAddress(gis_site, "Municipality"),
                 checkAddress(gis_site, "County"),
                 checkAddress(gis_site, "State"),
                 checkAddress(gis_site, "ZIP_Code"),
                 checkAddress(gis_site, "Parcel_ID"),
                 gis_site["Point"]["coordinates"].split(",")[0],
                 gis_site["Point"]["coordinates"].split(",")[1],
                 checkAddress(gis_site, "Owner"),
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Market_Value__Land_"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Market_Value__Land_"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Market_Value__Building_"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Market_Value__Building_"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Acreage"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Acreage"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Land_Cover"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Land_Cover"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Elevation_Ft_"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Elevation_Ft_"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 checkAddress(gis_site, "Mailing_Address_1"),
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Legal_Description_1"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Legal_Description_1"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "School_District"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "School_District"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Updated"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Updated"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Place"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Place"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Transfer_Date"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Transfer_Date"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Owner_Occupied"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Owner_Occupied"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Legal_Description_2"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Legal_Description_2"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Legal_Description_3"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Legal_Description_3"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Land_Use_Code"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Land_Use_Code"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "USPS_Residential"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "USPS_Residential"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Mailing_Address_2"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Mailing_Address_2"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Sale_Price"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Sale_Price"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Year_Built"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Year_Built"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Num_Buildings"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Num_Buildings"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Mailing_Name"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Mailing_Name"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else "",
                 [t for t in gis_site["ExtendedData"]["SchemaData"]["SimpleData"] if
                  t["@name"] == "Building_Area_SqFt"][0]["#text"] if "#text" in [t for t in gis_site["ExtendedData"][
                     "SchemaData"]["SimpleData"] if t["@name"] == "Building_Area_SqFt"][
                     0].keys() else "" if "ExtendedData" in gis_site.keys() else ""
                 ]
            )
        else:
            print("no")
        ids.append(max(ids) + 1)

workbook.save("newtest.xlsx")

pprint(potential_sites)
#print(as_dict["kml"]["Folder"]["Document"]["Folder"][1]['name'])