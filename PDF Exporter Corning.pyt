import arcpy
import arcpy.da
import arcpy.management
import arcpy.cartography
import arcpy.mp
import os
import time
import datetime
import tempfile
import webbrowser

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Export Tools"
        self.alias = "ExportTools"

        # List of tool classes associated with this toolbox
        self.tools = [PDFExporter]

class PDFExporter(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "PDF Exporter"
        self.description = "Exports a compled PDF to a specified location or a 'PDF' folder in the project directory"
        self.canRunInBackground = False
        self.layouts = [lay.name for lay in arcpy.mp.ArcGISProject("CURRENT").listLayouts()]
        self.pd_order = \
                [
                ["Cover"],
                ["Overview_Frames"],
                ["General Notes 1"],
                ["Legend"],
                ["BOM"],
                ["PlanView_PD"],
                ["Planview Typical"],
                ["TYP 1"],
                ["TYP 2"],
                ["TYP 3"],
                ["TCP 1"],
                ["TCP 2"],
                ["TCP 3"]
                ]
        if "Cover_PD" in self.layouts:
            self.pd_order[0] = ["Cover_PD"]
        self.pd_check = (True if set([i[0] for i in self.pd_order]).issubset(self.layouts) else False)
        self.cd_order = \
                [
                ["Cover"],
				#["BOM"],
                ["Overview_Frames"],
                ["Overview_Schematic"],
                ["Legend"],
                ["PlanView_CD"],
                ["PlanViewDetails_CD"],
                ["SpliceDetails"]
                ]
        if "Cover_CD" in self.layouts:
            self.cd_order[0] = ["Cover_CD"]
        self.cd_check = (True if set([i[0] for i in self.cd_order]).issubset(self.layouts) else False)
        self.cd_MakeReady_order = \
                [
                ["Cover_CD"],
				#["BOM"],
                ["Overview_Frames"],
                ["Overview_Schematic"],
                ["Legend"],
                ["PlanView_CD"],
                ["PlanViewDetails_CD"],
                ["SpliceDetails"],
                ["MakeReadySheet"]                
                ]
        if "Cover_CD" in self.layouts:
            self.cd_order[0] = ["Cover_CD"]
        self.cd_check = (True if set([i[0] for i in self.cd_order]).issubset(self.layouts) else False)

    def getParameterInfo(self):

        # Constants
        pd_cd = arcpy.Parameter(
            name="pd_cd",
            displayName="PD or CD",
            direction="Input",
            datatype="GPString",
            parameterType="Required"
        )
        pd_cd.value = "CD"
        pd_cd.filter.list = ["PD", "CD",  "CD_MakeReady", "Custom"]

        layout_order = arcpy.Parameter(
            name="layout_order",
            displayName="Layout Order",
            direction="Input",
            datatype="GPValueTable",
            parameterType="Required"
        )
        layout_order.columns = [["GPString", "Layout"]]
        layout_order.filters[0].list = self.layouts
        
        interlaced = arcpy.Parameter(
            name="interlaced",
            displayName="Interlaced Layouts",
            direction="Input",
            datatype="GPValueTable",
            parameterType="Optional"
        )
        interlaced.columns = [["GPString", "Interlaced Layout"]]
        interlaced.filters[0].list = self.layouts

        page_offset = arcpy.Parameter(
            name="page_offset",
            displayName="Page Offset (Pages before PlanView)",
            direction="Input",
            datatype="GPLong",
            parameterType="Optional"
        )
        page_offset.value = 5

        export_location = arcpy.Parameter(
            name = "export_location",
            displayName="Export Location",
            direction="Input",
            datatype="DEFile",
            parameterType="Required"
        )
        prj = arcpy.mp.ArcGISProject("CURRENT")
        if not os.path.exists(f"{prj.homeFolder}{os.sep}Export{os.sep}PDF"):
            os.makedirs(f"{prj.homeFolder}{os.sep}Export{os.sep}PDF")
        export_location.value = f"{prj.homeFolder}{os.sep}Export{os.sep}PDF"
        
        file_name = arcpy.Parameter(
            name="file_name",
            displayName="File Name",
            direction="Input",
            datatype="GPString",
            parameterType="Required"
        )
        file_name.value = prj.filePath.split(os.sep)[-1].replace('.aprx', '')
        
        quality = arcpy.Parameter(
            name="quality",
            displayName="Quality",
            direction="Input",
            datatype="GPString",
            parameterType="Required"
        )
        quality.value = "PRODUCTION"
        quality.filter.list = ["TEST", "DRAFT", "PRODUCTION"]

        pdf_open = arcpy.Parameter(
            name="pdf_open",
            displayName="Open PDF?",
            direction="Input",
            datatype="GPBoolean",
            parameterType="Optional"
        )
        pdf_open.value = True

        return [layout_order, export_location, file_name, quality, interlaced, page_offset, pdf_open, pd_cd]
    
    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True
    
    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        if parameters[7].value == "PD" and self.pd_check:
            parameters[0].value = self.pd_order
            if parameters[2].value.__contains__("_PD") or parameters[2].value.__contains__("_CD"):
                parameters[2].value = parameters[2].value.replace("_CD", "_PD")
            else:
                parameters[2].value = parameters[2].value.split('.')[0] + "_PD.pdf"
            if "Planview Typical" in self.layouts:
                parameters[4].value = [["Planview Typical"]]
        elif parameters[7].value == "CD" and self.cd_check:
            parameters[0].value = self.cd_order
            parameters[4].value = None
            if parameters[2].value.__contains__("_PD") or parameters[2].value.__contains__("_CD"):
                parameters[2].value = parameters[2].value.replace("_PD", "_CD")
            else:
                parameters[2].value = f"V1_Corning_Lima_{parameters[2].value}_CD_{datetime.datetime.today().strftime('%d_%m_%Y')}.pdf"
        elif parameters[7].value == "CD_MakeReady" and self.cd_check:
            parameters[0].value = self.cd_MakeReady_order
            parameters[4].value = None
            if parameters[2].value.__contains__("_PD") or parameters[2].value.__contains__("_CD"):
                parameters[2].value = parameters[2].value.replace("_PD", "_CD")
            else:
                parameters[2].value = f"V1_Corning_Lima_{parameters[2].value}_CD_{datetime.datetime.today().strftime('%d_%m_%Y')}.pdf"
        return
    
    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool parameter.
        This method is called after internal validation."""
        
        return
    
    def execute(self, parameters, messages):
        """The source code of the tool."""
        
        ## Set Export Quality
        quality = str(parameters[3].value)
        
        if quality == "TEST":
            resolution = 72
            image_quality = "FASTEST"
        elif quality == "DRAFT":
            resolution = 150
            image_quality = "NORMAL"
        elif quality == "PRODUCTION":
            resolution = 300
            image_quality = "BEST"
        else:
            msg("Quality must be TEST, DRAFT, or PRODUCTION", "error")
            return

        ## Initialize variables from parameters
        layout_order = [lay.strip("'") for lay in parameters[0].valueAsText.split(";")]
        if parameters[4].valueAsText:
            interlaced = [lay.strip("'") for lay in parameters[4].valueAsText.split(";")]
        else:
            interlaced = []
        offset = int(parameters[5].value)+1
        export_root = parameters[1].valueAsText
        filename = parameters[2].valueAsText
        export_loc = f"{export_root}{os.sep}{filename}"
        ## Timestamp Format
        timestamp = datetime.datetime.now().strftime('%b %d %Y, %I %M %S %p')
        pdf_open = parameters[6].value

        archive_loc = f"{export_root}{os.sep}_Archive{os.sep}{filename.split('.')[0]}"
        archive_file = f"{filename}_old({timestamp})"

        ## Initalize Archive folder
        if not os.path.exists(f"{export_root}{os.sep}_Archive"):
            os.mkdir(f"{export_root}{os.sep}_Archive")

        ## Archive and Datestamp Previous export
        if os.path.exists(export_loc) and os.path.exists(archive_loc):
            os.rename(export_loc, f"{archive_loc}{os.sep}{archive_file}")
        elif os.path.exists(export_loc) and not os.path.exists(archive_loc):
            os.mkdir(archive_loc)
            os.rename(export_loc, f"{archive_loc}{os.sep}{archive_file}")
        pdf = arcpy.mp.PDFDocumentCreate(export_loc)
        temp = tempfile.gettempdir()


        ## Print and Interlace
        prj = arcpy.mp.ArcGISProject("CURRENT")
        layouts = dict([[lay.name, lay] for lay in  prj.listLayouts()])
        arcpy.SetProgressor("step", "Exporting PDF", 0, len(layout_order), 1)
        page_dict = {}
        for layout in layout_order:
            arcpy.SetProgressorLabel("Processing layout: " + layout)
            msg("Processing layout: " + layout)
            if layouts[layout].mapSeries is not None:
                ## Set mapseries object and refresh
                ms = layouts[layout].mapSeries
                ms.refresh()
                ## Interlace the interlace pages
                if layout in interlaced:
                    idx = 0
                    for pnum in range(1, ms.pageCount + 1):
                        ms.currentPageNumber = pnum
                        parent_page = ms.pageRow.PageFinal.split(".")[0]
                        inter_page = ms.exportToPDF(
                                                    page_range_type= "CURRENT", 
                                                    out_pdf = f"{temp}{os.sep}{layout}.pdf", 
                                                    resolution=resolution, 
                                                    image_quality=image_quality
                                                    )
                        for page in page_dict:
                            page_dict[page] += 1
                        pdf.insertPages(inter_page, page_dict[parent_page]) # + offset)
                        arcpy.AddMessage(f"Index: {page_dict[parent_page]}")
                        arcpy.AddMessage(f"\tInterlaced {ms.pageRow.PageFinal} with {parent_page}")
                ## Append the Plan Pages
                else:
                    for pnum in range(1, ms.pageCount + 1):
                        ms.currentPageNumber = pnum
                        try:
                            page_dict[ms.pageRow.PageFinal] = offset
                            offset +=1
                        except:
                            pass
                    arcpy.AddMessage(f"{page_dict}")
                    pages = ms.exportToPDF(out_pdf = f"{temp}{os.sep}{layout}.pdf", resolution=resolution, image_quality=image_quality)
                    pdf.appendPages(pages)
            ## Append Other Pages
            else:
                lay = layouts[layout]
                pages = lay.exportToPDF(f"{temp}{os.sep}{layout}.pdf")
                pdf.appendPages(pages)
        pdf.saveAndClose()
        if pdf_open:
            os.startfile(export_loc)
        webbrowser.open("https://i.giphy.com/media/m2Q7FEc0bEr4I/giphy.webp")
        return
    

def msg(message="",level="message"):
    """
    Uses print() and the arcpy.AddMessage() function to print a message.
    
    Levels: 'message', 'warning', 'error'
    
    'message' is default
        
    usage: msg("<message>")
           msg("<warning_message>", "warning")
           msg("<error_message>", "error")
    
    Calling with no arguments will print a blank line with level message.
    
    Cautions: Fails silently if unable to print to the console or Arc Messagebox
    """
    message = str(message)
    level = str(level).lower()
    level = ("message" if level not in ["message", "warning", "error"] else level)
    try:
        # Message
        if level == "message":
            print(message)
            arcpy.AddMessage(message)
        # Warning
        elif level == "warning":
            print(f"WARNING: {message}")
            arcpy.AddWarning(message)
        # Error
        elif level == "error":
            print(f"ERROR: {message}")
            arcpy.AddError(message)
    except Exception as e:
        return e