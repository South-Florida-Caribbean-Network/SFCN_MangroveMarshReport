###################################
# SFCN_MangroveMarsh_Tables_Figures.py
###################################
# Description:  Routine to Summarize Mangrove Marsh Monitoring data including output tables and figures for annual reporting.

# Code performs the following routines:
# 1) Summaries the Event Average Distance from Ground Truth values.  This includes Averge, Standard Error, Confidence Interval for defined %, and Maximum and Minimum values
# by event/segment replicating the 'Mangrove-Marsh Ecotone Monitoring' SOP8-1 summary table. Confidence Intervals are defined using a Student T Distribution with Student T defiend as: np.abs(t.ppf((1 - confidence) / 2, dof)).

# 2) Quick QAQC of the data 

# 3) Summarize via a CrossTab/Pivot Table the Absolute Vegetation by Region, Location Name (i.e. Point on Segment), by Community Type, and by Taxon (Scale is Point - single value).
# Routine replicates the 'Mangrove-Marsh Ecotone Monitoring' SOP8-2 summary table.

# 4) Calculates the Absolute Cover By Region, By Community, By Strata - data is from table  'tbl_MarkerData'. 

# 5) Create 'Mangrove-Marsh Ecotone Monitoring' SOP8-3 summary Figure.

# Output:
# An excel spreadsheet with:
## Tables SOP8-1 (i.e. Average Distance from Ground Truth Values)
## QAQC Table confirming that 1) Relative Cover of Strata = 100% & 
## Tables SOP8-2 By region (three total) with the the Absolute Vegetation by Location Name (i.e. Point on Segment), by Community Type, and by Taxon 
# A PDF file withe Figures SOP8-3 by region (three total) with the Absolute Cover by Community and Strata.

# Dependencies:
# Python version 3.12
# Pandas
# Scipyby

# Date Created: March 2023 - Kirk Sherrill
# Date Updated: March 2026 - Caitlin Andrews

from datetime import date, datetime
import logging
import os
import pandas as pd
import numpy as np
from scipy.stats import t

import matplotlib as mp
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages #Imort PdfPages from MatplotLib

import pyodbc
pyodbc.pooling = False  #So you can close the pydobx connection

#######################################
# 1. Define Input Parameters:
#######################################

#Directory Information
current_directory = os.getcwd()
outputDir = current_directory  #Output Directory
workspace = current_directory  # Workspace Folder
monitoringYear = 2020  #Monitoring Year of Mangrove Marsh data being processing

#Mangrove Marsh Access Database and location
inDB = r'Z:\Files\Vital_Signs\Mangrove_Marsh_Ecotone\data\vegetation_databases\SFCN_Mangrove_Marsh_Ecotone_tabular_20260109.mdb'

#Define Output Names 
dateString = date.today().strftime("%Y%m%d")
outName = "MangroveMarsh_AnnualTablesFigs_" + str(monitoringYear) + "_" + dateString  # Name given to the exported pre-processed
outPDF = outputDir + "\\" + outName + ".pdf"

# Logging info
logFileName = workspace + "\\" + outName + "_logfile.txt"
logging.basicConfig(
    filename=logFileName,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

#Confidence Interval
confidence = 0.95

# Function to Get the Date/Time
def timeFun():
    b = datetime.now()
    messageTime = b.isoformat()
    return messageTime

#Connect to Access DB and perform defined query - return query in a dataframe
def connect_to_AccessDB(query, inDB):

    try:
        connStr = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' r'DBQ=' + inDB + r';')
        cnxn = pyodbc.connect(connStr)
        queryDf = pd.read_sql(query, cnxn)
        cnxn.close()
        return "success function", queryDf

    except:
        scriptMsg = f"Error function:  connect_to_AccessDB - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        logging.exception("WARNING Script Failed - connect_to_AccessDB")
        return "failed function"

def main():
    try:

        ########################
        #Functions for Table 8-1
        ########################
        
        #Pull Data from Database for table 'tbl_MarkerData'
        outVal = defineRecords_MarkerData()
        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_MarkerData - Failed - Exiting Script")
            exit()

        print("COMPLETE: Function defineRecords_MarkerData")
        outDF = outVal[1]

        #Summarize Data from tbl_MarkerData into format for Table 8-1 in SFCN Mangrove Marsh SOP.
        outVal = SummarizeFigure8_1(outDF)
        if outVal[0].lower() != "success function":
            print("WARNING - Function SummarizeFigure8_1 - Failed - Exiting Script")
            exit()
        
        scriptMsg = f"COMPLETE: Function SummarizeFigure8_1 - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

        ########################
        # Functions for QAQC - Relative Cover By Strata and Within Strata
        ########################    
        QAQC_RelativeCover1()

        scriptMsg = f"COMPLETE - QAQC Functions - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

        ########################
        # Functions for Table 8-2  - Absolute Cover Species Data by Transect and Point By Region
        ########################

        #Summarize via a CrossTab/Pivot Table the Absolute Vegetation by Location Name (i.e. Point on Segment), by Community Type and Vegetation Type (Scale is Point - single value)
        outVal = defineRecords_VegCoverByPointAbsolute()
        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_VegCoverByPointAbsolute - Failed - Exiting Script")
            exit()

        scriptMsg = f"COMPLETE: SFCN_MangroveMash_Table_8-2 - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

        ########################
        # Functions for Figure 8-3  - Calculate the Absolute Cover By Region, By Community, By Strata - data is from table  'tbl_MarkerData'
        ########################

        #Pull Stratum Cover Data by point in 'tbl_MarkerData'
        outVal = defineRecords_CoverByStratum()
        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_CoverByStratum - Failed - Exiting Script")
            exit()

        outDF = outVal[1]

        #Figures for Marker Point Stratum By Region - Stacked Top Marsh, Botom Mangrove
        outVal = figure_CoverByStratum(outDF)
        if outVal.lower() != "success function":
            print("WARNING - Function figure_CoverByStratum - Failed - Exiting Script")
            exit()

        scriptMsg = f"COMPLETE - SFCN_MangroveMash_Tables_Figures_8-3"
        print(scriptMsg)
        logging.info(scriptMsg)

        scriptMsg = f"ALL COMPLETE - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

    except:

        logging.exception(f"WARNING Script Failed - {timeFun()}")

#Summarize Mangrove Marsh Ecotone Values - Average Distance, Standard Error, Lower 95% Confidence Limit, Upper 95% Confidence Limit, Max and Min Values
## Calculate the Confidence Interval Upper 95% and Lower 95% using Students T Distribution
### Student T Distribution is defined as t_crit = np.abs(t.ppf((1-confidence)/2,dof))
### CI Upper and lower is: (Mean +/- t_crit * SE)  
#Output DataFrame with the summary values by Segment
def SummarizeFigure8_1(inDF):
    try:
        # Calculate Average, Standard Error, and count of Distance 
        outDf_8pt1 = (
            inDF.groupby(['Event_Group_ID', 'Region', 'Segment'], as_index=False)
            .agg(
                AverageDist_M=('Distance', 'mean'),
                StandardError=('Distance', 'sem'),
                RecCount=('Distance', 'count'),
                MinDifference=('Distance', 'min'),
                MaxDifference=('Distance','max')
            )
            .assign(
                DOF=lambda d: d.RecCount - 1,
                t_crit=lambda d: np.abs(t.ppf((1 + confidence) / 2, d.DOF)),
                LowerCI_95=lambda d: d.AverageDist_M - d.t_crit * d.StandardError,
                UpperCI_95=lambda d: d.AverageDist_M + d.t_crit * d.StandardError,
                SortField=lambda d: pd.to_numeric(
                    d.Segment.str.extract(r'(\d+)')[0],
                    errors='coerce',
                    downcast='integer'
                )
            )
            .sort_values('SortField')
            .set_index('SortField')
            .drop(columns=['DOF', 't_crit'])
            )
        
        # Export Table to Excel
        dateString = date.today().strftime("%Y%m%d")
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_{dateString}.xlsx")
        outDf_8pt1.to_excel(outFull, sheet_name = 'SOP8-1', index=False)

        scriptMsg = f"EXPORTED Table 8-1 to: {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        return "success function", outDf_8pt1

    except:
        print(f"Error on SummarizeFigure8_1 Function - {timeFun()}")
        logging.exception("Error in SummarizeFigure8_1")
        return "Failed function - 'SummarizeFigure8_1'"

# QAQC the Vegetation Data
## Check for correct relative cover (100%) across strata within community type / location
## Check for correct relative cover (100%) within strata / community type / location
def QAQC_RelativeCover1():
    try:
        # get data from access
        inQuery = """
        SELECT 
            Region,
            Location_Name,
            Segment,
            MangroveSide_Cover_Overall,
            MangroveSide_Cover_Tree,
            MangroveSide_Cover_Shrub,
            MangroveSide_Cover_Herb,
            MarshSide_Cover_Overall,
            MarshSide_Cover_Tree,
            MarshSide_Cover_Shrub,
            MarshSide_Cover_Herb
        FROM 
            (tbl_MarkerData
            INNER JOIN tbl_Events 
                ON tbl_MarkerData.Event_ID = tbl_Events.Event_ID)
            INNER JOIN tbl_Locations 
                ON tbl_Events.Location_ID = tbl_Locations.Location_ID
        """
        outVal = connect_to_AccessDB(inQuery, inDB)

        if outVal[0].lower() != "success function":
            print(f"QAQC_RelativeCover1 Database Query - FAILED - Exiting Script - {timeFun()}")
            return

        outDF = outVal[1]

        # sum relative cover and check
        outDF["sum_mangrove"] = outDF[
            ["MangroveSide_Cover_Tree", "MangroveSide_Cover_Shrub", "MangroveSide_Cover_Herb"]
        ].sum(axis=1)

        outDF["sum_marsh"] = outDF[
            ["MarshSide_Cover_Tree", "MarshSide_Cover_Shrub", "MarshSide_Cover_Herb"]
        ].sum(axis=1)

        outDF["mangrove_is_100"] = outDF["sum_mangrove"] == 100
        outDF["marsh_is_100"] = outDF["sum_marsh"] == 100

        # Append DataFrame to existing excel file
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_{dateString}.xlsx")

        with pd.ExcelWriter(outFull, mode='a', engine="openpyxl") as writer:
            outDF.to_excel(writer, sheet_name='QAQC-1-RelCov', index=False)

        scriptMsg = f"EXPORTED Table QAQC-1-RelCover to {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

    except:
        print(f"Error on QAQC_RelativeCover1 Function - {timeFun()}")
        logging.exception("Error in QAQC_RelativeCover1")
        return "Failed function - 'QAQC_RelativeCover1'"

# Summarize via a CrossTab/Pivot Table the Absolute Cover By Region, Community, Strata and Taxon across point locations
#Export By Region
def defineRecords_VegCoverByPointAbsolute():

    try:
        dateString = date.today().strftime("%Y%m%d")
        # Process By Region
        regionlList = ['Turner River', 'Shark Slough', 'Taylor Slough']
        for count, region in enumerate(regionlList):
            inQuery = "TRANSFORM Sum(IIf([VegetationType]='Tree' And [CommunityType]='Mangrove',([MangroveSide_Cover_Overall])*([MangroveSide_Cover_Tree]/100)*([PercentCover]/100),IIf([VegetationType]='Shrub' And"\
                " [CommunityType]='Mangrove',([MangroveSide_Cover_Overall])*([MangroveSide_Cover_Shrub]/100)*([PercentCover]/100),IIf([VegetationType]='Herb' And"\
                " [CommunityType]='Mangrove',([MangroveSide_Cover_Overall])*([MangroveSide_Cover_Herb]/100)*([PercentCover]/100),IIf([VegetationType]='Tree' And"\
                " [CommunityType]='Marsh',([MarshSide_Cover_Overall])*([MarshSide_Cover_Tree]/100)*([PercentCover]/100),IIf([VegetationType]='Shrub' And"\
                " [CommunityType]='Marsh',([MarshSide_Cover_Overall])*([MarshSide_Cover_Shrub]/100)*([PercentCover]/100),IIf([VegetationType]='Herb' And"\
                " [CommunityType]='Marsh',([MarshSide_Cover_Overall])*([MarshSide_Cover_Herb]/100)*([PercentCover]/100),-999))))))) AS AbsolutePercCover"\
                " SELECT tbl_Locations.Region, tbl_MarkerData_Vegetation.CommunityType, tbl_MarkerData_Vegetation.VegetationType, tlu_Vegetation.ScientificName"\
                " FROM tbl_Locations INNER JOIN (((tbl_Event_Group INNER JOIN tbl_Events ON (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID) AND"\
                " (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID)) INNER JOIN tbl_MarkerData ON (tbl_Events.Event_ID = tbl_MarkerData.Event_ID) AND"\
                " (tbl_Events.Event_ID = tbl_MarkerData.Event_ID)) INNER JOIN (tbl_MarkerData_Vegetation LEFT JOIN tlu_Vegetation ON tbl_MarkerData_Vegetation.SpeciesCode = tlu_Vegetation.SpeciesCode)"\
                " ON (tbl_MarkerData.Point_ID = tbl_MarkerData_Vegetation.Point_ID) AND (tbl_MarkerData.Point_ID = tbl_MarkerData_Vegetation.Point_ID)) ON tbl_Locations.Location_ID"\
                " = tbl_Events.Location_ID WHERE ((Not (tbl_MarkerData_Vegetation.PercentCover) Is Null) AND ((tbl_Events.Event_Type)='Marker Visit') AND ((tbl_Locations.Region)= '" + region + "'))"\
                " GROUP BY tbl_Locations.Region, tbl_MarkerData_Vegetation.CommunityType, tbl_MarkerData_Vegetation.VegetationType, tlu_Vegetation.ScientificName"\
                " ORDER BY tbl_MarkerData_Vegetation.CommunityType, tbl_MarkerData_Vegetation.VegetationType, tlu_Vegetation.ScientificName"\
                " PIVOT tbl_Locations.Location_Name;"
            outVal = connect_to_AccessDB(inQuery, inDB)

            if outVal[0].lower() != "success function":
                print(f"WARNING - Function defineRecords_VegCoverBySegment - Failed - Exiting Script- {timeFun()}")
                exit()

            outDF = outVal[1]
            print(f"Success: defineRecords_VegCoverBySegment - {timeFun()}")

            # Append DataFrame to existing excel file
            outFull = os.path.join(outputDir, f"MangroveMarsh_Export_{dateString}.xlsx")

            with pd.ExcelWriter(outFull, mode='a', engine="openpyxl") as writer:
                outDF.to_excel(writer, sheet_name='SOP8-2-AbsCov-' + region, index=False)

            scriptMsg = f"EXPORTED Table 8-2-AbsCov - {region} to {outFull} - {timeFun()}"
            print(scriptMsg)
            logging.info(scriptMsg)

        return "success function", outDF
    except:
        print(f"Error on defineRecords_VegCoverByPointAbsolute - {timeFun()}")
        logging.exception("WARNING Script Failed - defineRecords_VegCoverByPointAbsolute")
        return "Failed function - 'defineRecords_VegCoverByPointAbsolute'"

#Create  Figures - Absolute Cover By Region, By Community, By Strata
def figure_CoverByStratum(inDF):
    try:

        #Open PDF to be copied to
        pdf = PdfPages(outPDF)

        # Process By Region
        regionlList = ['Turner River', 'Shark Slough', 'Taylor Slough']
        for count, region in enumerate(regionlList):

            #Subset By Region
            outDFSub = inDF[inDF['Region'] == region]

            ###########################
            #Marsh DataFrame and Figure
            ###########################
            #Subset to marshDF fields
            marshDF = outDFSub.loc[:,('Location_Name', 'MarshSide_Cover_Overall', 'AbsCover_Marsh_Tree', 'AbsCover_Marsh_Shrub','AbsCover_Marsh_Herb')]
            #Rename Fields
            marshDF.rename(columns={"AbsCover_Marsh_Tree": "Tree", "AbsCover_Marsh_Shrub": "Shrub", "AbsCover_Marsh_Herb": "Herb"}, inplace=True)

            #Set Index
            marshDF.set_index('Location_Name', inplace=True)

            #Subset to Mangrove fields
            mangroveDF = outDFSub.loc[:, ('Location_Name', 'MangroveSide_Cover_Overall', 'AbsCover_Mangrove_Tree','AbsCover_Mangrove_Shrub', 'AbsCover_Mangrove_Herb')]
            # Rename Fields
            mangroveDF.rename(columns={"AbsCover_Mangrove_Tree": "Tree", "AbsCover_Mangrove_Shrub": "Shrub", "AbsCover_Mangrove_Herb": "Herb"}, inplace=True)

            #Set Index
            mangroveDF.set_index('Location_Name', inplace=True)

            ####################
            #Create the Figures:
            ####################
            plt.figure(figsize=(8, 6))
            ax1 = plt.subplot(2, 1, 1)
            marshDF.plot.bar(stacked=True, title="Marsh Side - " + region, xlabel="Marker Points", ylabel="Absolute Percent Cover (%)", color={'Shrub': 'red', 'Herb': 'turquoise', 'Tree': 'orange'}, ax=ax1)
            lgd = plt.legend(['Tree', 'Shrub', 'Herb'], loc='center left', bbox_to_anchor=(1, 0.5))
            plt.grid(axis='y')
            plt.ylim(0, 100)
            plt.tight_layout(pad=0.4)

            ax2 = plt.subplot(2, 1, 2)
            mangroveDF.plot.bar(stacked=True, title="Mangrove Side - " + region, xlabel="Marker Points", ylabel="Absolute Percent Cover (%)", color={'Shrub': 'red', 'Herb': 'turquoise', 'Tree': 'orange'}, ax=ax2)
            lgd = plt.legend(['Tree', 'Shrub', 'Herb'], loc='center left', bbox_to_anchor=(1, 0.5))
            plt.grid(axis='y')
            plt.ylim(0, 100)
            plt.tight_layout(pad=0.4)

            figure = mp.pyplot.gcf()
            pdf.savefig(figure)

            scriptMsg = f"EXPORTED Figures Region: {region} - {timeFun()}"
            print(scriptMsg)
            logging.info(scriptMsg)

        pdf.close()

        print(f"Success: figure_CoverByStratum - {timeFun()}")
        return "success function"

    except:
        print(f"Error on figure_CoverByStratum Function - {timeFun()}")
        logging.exception("WARNING Script Failed - figure_CoverByStratum")
        return "Failed function - 'figure_CoverByStratum'"

#Calculate Absolute Cover By Stratum and Community type in table 'tbl_MarkerData'
def defineRecords_CoverByStratum():
    try:
        inQuery = "SELECT tbl_Locations.Location_ID, tbl_Locations.Order_ID,  tbl_Locations.Region, tbl_Locations.Location_Name,  tbl_Events.Event_ID, tbl_Event_Group.Start_Date, tbl_MarkerData.MangroveSide_Cover_Overall,"\
                " tbl_MarkerData.MangroveSide_Cover_Tree, tbl_MarkerData.MangroveSide_Cover_Shrub, tbl_MarkerData.MangroveSide_Cover_Herb, [MangroveSide_Cover_Overall]*([MangroveSide_Cover_Tree]/100)"\
                " AS AbsCover_Mangrove_Tree, [MangroveSide_Cover_Overall]*([MangroveSide_Cover_Shrub]/100) AS AbsCover_Mangrove_Shrub, [MangroveSide_Cover_Overall]*([MangroveSide_Cover_Herb]/100)"\
                " AS AbsCover_Mangrove_Herb, tbl_MarkerData.MarshSide_Cover_Overall, tbl_MarkerData.MarshSide_Cover_Tree, tbl_MarkerData.MarshSide_Cover_Shrub, tbl_MarkerData.MarshSide_Cover_Herb,"\
                " [MarshSide_Cover_Overall]*([MarshSide_Cover_Tree]/100) AS AbsCover_Marsh_Tree, [MarshSide_Cover_Overall]*([MarshSide_Cover_Shrub]/100) AS AbsCover_Marsh_Shrub,"\
                " [MarshSide_Cover_Overall]*([MarshSide_Cover_Herb]/100) AS AbsCover_Marsh_Herb"\
                " FROM tbl_Locations INNER JOIN ((tbl_Event_Group INNER JOIN tbl_Events ON (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID) AND (tbl_Event_Group.Event_Group_ID"\
                " = tbl_Events.Event_Group_ID)) INNER JOIN tbl_MarkerData ON (tbl_Events.Event_ID = tbl_MarkerData.Event_ID) AND (tbl_Events.Event_ID = tbl_MarkerData.Event_ID))"\
                " ON tbl_Locations.Location_ID = tbl_Events.Location_ID WHERE (((tbl_Events.Event_Type)='Marker Visit')) ORDER BY tbl_Locations.Order_ID, tbl_Locations.Location_Name,"\
                " tbl_Event_Group.Start_Date;"
        outVal = connect_to_AccessDB(inQuery, inDB)

        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_CoverByStratum - Failed - Exiting Script -  {timeFun()}")
            exit()
        
        outDF = outVal[1]
        print(f"Success:  defineRecords_CoverByStratum - {timeFun()}")
        return "success function", outDF

    except:
        print(f"Error on defineRecords_CoverByStratum Function - {timeFun()}")
        logging.exception("WARNING Script Failed - defineRecords_CoverByStratum")
        return "Failed function - 'defineRecords_CoverByStratum'"

#Extract Mangrove Marsh Distance Records table 'tbl_MarkerData' where Event Type = 'Marker Visit'
def defineRecords_MarkerData():
    try:
        inQuery = "SELECT tbl_Event_Group.Event_Group_ID, tbl_Event_Group.Event_Group_Name, tbl_Event_Group.Start_Date, tbl_Event_Group.End_Date, tbl_Event_Group.Assessment, tbl_Events.Event_Type,"\
                " tbl_Events.Location_ID, tbl_Locations.Region,tbl_Locations.Segment, tbl_Locations.Location_Name, tbl_MarkerData.Distance, tbl_MarkerData.Method"\
                " FROM tbl_Locations INNER JOIN ((tbl_Event_Group INNER JOIN tbl_Events ON (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID) AND (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID))"\
                " INNER JOIN tbl_MarkerData ON (tbl_Events.Event_ID = tbl_MarkerData.Event_ID) AND (tbl_Events.Event_ID = tbl_MarkerData.Event_ID)) ON tbl_Locations.Location_ID = tbl_Events.Location_ID"\
                " WHERE tbl_Events.Event_Type = 'Marker Visit' ORDER BY tbl_Locations.Segment, tbl_Locations.Location_Name, tbl_Events.Event_Type;"\

        outVal = connect_to_AccessDB(inQuery, inDB)
        if outVal[0].lower() != "success function":
            print(f"WARNING - Function defineRecords_MarkerData - Failed - Exiting Script - {timeFun()}")
            exit()
            
        outDF = outVal[1]
        print(f"Success:  defineRecords_MarkerData - {timeFun()}")
        return "success function", outDF

    except:
        print(f"Error on defineRecords_MarkderData Function - {timeFun()}")
        logging.exception("WARNING Script Failed - defineRecords_MarkderData")
        return "Failed function - 'defineRecords_MarkderData'"

if __name__ == '__main__':

    # Checks ---------------------------------------------
    if os.path.exists(workspace):
        pass
    else:
        os.makedirs(workspace)

    #Check for logfile

    if os.path.exists(logFileName):
        pass
    else:
        logFile = open(logFileName, "w")    #Creating index file if it doesn't exist
        logFile.close()

    # Analyses routine ---------------------------------------------------------
    main()