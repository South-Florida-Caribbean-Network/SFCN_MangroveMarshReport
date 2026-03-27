###################################
# SFCN_MangroveMarsh_Tables_Figures.py
###################################
# Description:  Routine to Summarize Mangrove Marsh Monitoring data including output tables and figures for annual reporting.

# Code performs the following routines:
# 1) Summaries the Event Average Distance from Ground Truth values.  This includes Averge, Standard Error, Confidence Interval for defined %, and Maximum and Minimum values
# by event/segment replicating the 'Mangrove-Marsh Ecotone Monitoring' SOP8-1 summary table. Confidence Intervals are defined using a Student T Distribution with Student T defiend as: np.abs(t.ppf((1 - confidence) / 2, dof)).

# 2) Quick QAQC of the data - Relative Cover checks

# 3) Summarize via a CrossTab/Pivot Table the Absolute Vegetation by Region, Location Name (i.e. Point on Segment),
# by Community Type, and by Taxon (Scale is Point - single value).
# Routine replicates the 'Mangrove-Marsh Ecotone Monitoring' SOP8-2 summary table.

# 4) Calculates the Absolute Cover By Region, By Community, By Strata - data is from table  'tbl_MarkerData'. 

# 5) Create 'Mangrove-Marsh Ecotone Monitoring' SOP8-3 summary Figure.

# Output:
# An excel spreadsheet with:
## Tables SOP8-1 (i.e. Average Distance from Ground Truth Values)
## QAQC Table confirming that 1) Relative Cover of Strata = 100% & 
## Tables SOP8-2 By region (three total) with the the Absolute Vegetation by Location Name (i.e. Point on Segment), by Community Type, and by Taxon 
# A PoutVal file withe Figures SOP8-3 by region (three total) with the Absolute Cover by Community and Strata.

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
#from matplotlib.backends.backend_poutVal import PoutValPages #Imort PoutValPages from MatplotLib

import pyodbc
pyodbc.pooling = False  #So you can close the pydobx connection

#######################################
# 1. Define Input Parameters:
#######################################

# --- Paths ---
current_directory = os.getcwd()
outputDir = current_directory  #Output Directory
workspace = current_directory  # Workspace Folder
monitoringYear = 2015  #Monitoring Year of Mangrove Marsh data being processing

#Mangrove Marsh Access Database and location
inDB = r'Z:\Files\Vital_Signs\Mangrove_Marsh_Ecotone\data\vegetation_databases\2009-2015_SFCN_Mangrove_Marsh_Ecotone.mdb'

# --- Output names ---
dateString = date.today().strftime("%Y%m%d")
outName = f"MangroveMarsh_AnnualTablesFigs_{str(monitoringYear)}_{dateString}"  # Name given to the exported pre-processed
outPoutVal = os.path.join(outputDir, f"{outName}.poutVal")

# --- Logging ---
logFileName = os.path.join(workspace, f"{outName}_logfile.txt")
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
        queryoutVal = pd.read_sql(query, cnxn)
        cnxn.close()
        return "success function", queryoutVal
    except:
        scriptMsg = f"Error function:  connect_to_AccessDB - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        logging.exception("WARNING Script Failed - connect_to_AccessDB")
        return "failed function"

# helper function for coomputing absolute cover
def compute_cover(row):
    ct = row["CommunityType"]
    vt = row["VegetationType"]
    try:
        if ct == "Mangrove":
            overall = row["MangroveSide_Cover_Overall"]
            specific = row[f"MangroveSide_Cover_{vt}"]
        elif ct == "Marsh":
            overall = row["MarshSide_Cover_Overall"]
            specific = row[f"MarshSide_Cover_{vt}"]
        else:
            return np.nan
        return overall * (specific / 100) * (row["PercentCover"] / 100)
    except KeyError:
        # handles unexpected VegetationType values
        return np.nan
    
def main():
    try:

        ########################
        #Functions for Table 8-1 - No Longer Active in this workflow
        ########################
        
        # #Pull Data from Database for table 'tbl_MarkerData'
        # outVal = defineRecords_MarkerData()
        # if outVal[0].lower() != "success function":
        #     print("WARNING - Function defineRecords_MarkerData - Failed - Exiting Script")
        #     exit()

        # print("COMPLETE: Function defineRecords_MarkerData")
        # outVal = outVal[1]

        # #Summarize Data from tbl_MarkerData into format for Table 8-1 in SFCN Mangrove Marsh SOP.
        # outVal = SummarizeFigure8_1(outVal)
        # if outVal[0].lower() != "success function":
        #     print("WARNING - Function SummarizeFigure8_1 - Failed - Exiting Script")
        #     exit()
        
        # scriptMsg = f"COMPLETE: Function SummarizeFigure8_1 - {timeFun()}"
        # print(scriptMsg)
        # logging.info(scriptMsg)

        ########################
        # Functions for QAQC - Relative Cover By Strata and Within Strata
        ########################    
        outVal = QAQC_RelCoverByStratum()
        print(outVal[0])
        if outVal[0].lower() != "s":
            print("WARNING - Function QAQC_RelCoverByStratum - Failed - Exiting Script")
            exit()
            
        outVal = QAQC_RelCoverByPoint()
        if outVal[0].lower() != "s":
           print("WARNING - Function QAQC_RelCoverByPoint - Failed - Exiting Script")
           exit()

        scriptMsg = f"COMPLETE - QAQC Functions - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

        ########################
        # Functions for Table 8-2  - Absolute Cover Species Data by Transect and Point By Region
        ########################

        # #Summarize via a CrossTab/Pivot Table the Absolute Vegetation by Location Name (i.e. Point on Segment), by Community Type and Vegetation Type (Scale is Point - single value)
        outVal = defineRecords_AbsCoverByPoint()
        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_AbsCoverByPoint - Failed - Exiting Script")
            exit()

        scriptMsg = f"COMPLETE: SFCN_MangroveMash_Table_8-2 - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

        # ########################
        # # Functions for Figure 8-3  - Calculate the Absolute Cover By Region, By Community, By Strata - data is from table  'tbl_MarkerData'
        # ########################

        # #Pull Stratum Cover Data by point in 'tbl_MarkerData'
        outVal = defineRecords_AbsCoverByStratum()
        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_AbsCoverByStratum - Failed - Exiting Script")
            exit()

        # outVal = outVal[1]

        # #Figures for Marker Point Stratum By Region - Stacked Top Marsh, Botom Mangrove
        # outVal = figure_AbsCoverByStratum(outVal)
        # if outVal.lower() != "success function":
        #     print("WARNING - Function figure_AbsCoverByStratum - Failed - Exiting Script")
        #     exit()

        # scriptMsg = f"COMPLETE - SFCN_MangroveMash_Tables_Figures_8-3"
        # print(scriptMsg)
        # logging.info(scriptMsg)

        # scriptMsg = f"ALL COMPLETE - {timeFun()}"
        # print(scriptMsg)
        # logging.info(scriptMsg)

    except:

        logging.exception(f"WARNING Script Failed - {timeFun()}")

## QAQC - Check for correct relative cover (100%) ACROSS strata for each location / community type
def QAQC_RelCoverByStratum():
    try:
        # get data from access
        inQuery = """
        SELECT 
            tbl_Events.Start_Date,
            geo_Locations.Loc_Name_Short AS Location_Name,
            Event_Type,
            MangroveSide_Cover_Tree,
            MangroveSide_Cover_Shrub,
            MangroveSide_Cover_Herb,
            MarshSide_Cover_Tree,
            MarshSide_Cover_Shrub,
            MarshSide_Cover_Herb
        FROM 
            (tbl_MarkerData
            INNER JOIN tbl_Events 
                ON tbl_MarkerData.Event_ID = tbl_Events.Event_ID)
            INNER JOIN geo_Locations 
                ON tbl_Events.Location_ID = geo_Locations.GlobalID
        ORDER BY Loc_Name_Short
        """
        outVal = connect_to_AccessDB(inQuery, inDB)

        if outVal[0].lower() != "success function":
            print(f"QAQC_RelCoverByStratum Database Query - FAILED - Exiting Script - {timeFun()}")
            return

        outVal = outVal[1]

        # sum relative cover and check
        outVal["sum_mangrove"] = outVal[
            ["MangroveSide_Cover_Tree", "MangroveSide_Cover_Shrub", "MangroveSide_Cover_Herb"]
        ].sum(axis=1)

        outVal["sum_marsh"] = outVal[
            ["MarshSide_Cover_Tree", "MarshSide_Cover_Shrub", "MarshSide_Cover_Herb"]
        ].sum(axis=1)

        outVal["mangrove_is_100"] = outVal["sum_mangrove"] == 100
        outVal["marsh_is_100"] = outVal["sum_marsh"] == 100

        # Write DataFrame to excel file
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_oldDB_{dateString}.xlsx")
        outVal.to_excel(outFull, sheet_name='QAQC_RelCovByStrata', index=False)

        scriptMsg = f"EXPORTED Table QAQC-RelCoverByStratum to {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        return "success function"

    except:
        print(f"Error on QAQC_RelativeCoverByStratum Function - {timeFun()}")
        logging.exception("Error in QAQC_RelativeCoverByStratum")
        return "Failed function - 'QAQC_RelativeCoverByStratum'"

## QAQC - Check for correct relative cover (100%) WITHIN strata for each location / community type
def QAQC_RelCoverByPoint():
    try:
        # get data from access
        inQuery = """
        SELECT 
            tbl_Events.Start_Date,
            geo_Locations.Loc_Name_Short AS Location_Name,
            Event_Type,
            CommunityType,
            VegetationType,
            SpeciesCode,
            PercentCover
        FROM 
            ((tbl_MarkerData_Vegetation
            INNER JOIN tbl_MarkerData
                ON tbl_MarkerData_Vegetation.Point_ID = tbl_MarkerData.Point_ID)
            INNER JOIN tbl_Events 
                ON tbl_MarkerData.Event_ID = tbl_Events.Event_ID)
            INNER JOIN geo_Locations 
                ON tbl_Events.Location_ID = geo_Locations.GlobalID
        ORDER BY Loc_Name_Short
        """
        outVal = connect_to_AccessDB(inQuery, inDB)

        if outVal[0].lower() != "success function":
            print(f"def QAQC_RelCoverByPoint(): Database Query - FAILED - Exiting Script - {timeFun()}")
            return

        outVal = outVal[1]

        # pivot and sum to see if each strata within a point = 100
        wide_outVal_count = outVal.pivot_table(
            index = ["Location_Name", "CommunityType", "VegetationType", "Event_Type"],
            columns = "SpeciesCode",
            values = "PercentCover",
            aggfunc = "count"
        )

        # pivot and sum to see if each strata within a point = 100
        wide_outVal_sum = outVal.pivot_table(
            index = ["Location_Name", "CommunityType", "VegetationType", "Event_Type"],
            columns = "SpeciesCode",
            values = "PercentCover",
            aggfunc = "sum"
        )

        wide_outVal_sum['sum_relcover'] = wide_outVal_sum.sum(axis=1)
        wide_outVal_sum["relcover_is_100"] = wide_outVal_sum["sum_relcover"] == 100

        # Append DataFrame to existing excel file
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_oldDB_{dateString}.xlsx")

        with pd.ExcelWriter(outFull, mode='a', engine="openpyxl") as writer:
            wide_outVal_count.to_excel(writer, sheet_name='QAQC_RelCoverByPoint_Count', index=True)

        with pd.ExcelWriter(outFull, mode='a', engine="openpyxl") as writer:
            wide_outVal_sum.to_excel(writer, sheet_name='QAQC_RelCoverByPoint_Sum', index=True)

        scriptMsg = f"EXPORTED Table QAQC_RelCoverByPoint to {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        return "success function"

    except:
        print(f"Error on QAQC_RelCoverByPoint Function - {timeFun()}")
        logging.exception("Error in QAQC_RelCoverByPoint")
        return "Failed function - 'QAQC_RelCoverByPoint'"

# Summarize via a CrossTab/Pivot Table the Absolute Cover By Region, Community, Strata and Taxon across point locations
#Export By Region
def defineRecords_AbsCoverByPoint():

    try:
        dateString = date.today().strftime("%Y%m%d")
        in_query = f"""
            SELECT 
                tbl_Events.Start_Date,
                geo_Locations.Loc_Name_Short AS Location_Name,
                Event_Type,
                CommunityType,
                VegetationType,
                SpeciesCode,
                PercentCover,
                MangroveSide_Cover_Overall,
                MangroveSide_Cover_Tree,
                MangroveSide_Cover_Shrub,
                MangroveSide_Cover_Herb,
                MarshSide_Cover_Overall,
                MarshSide_Cover_Tree,
                MarshSide_Cover_Shrub,
                MarshSide_Cover_Herb

            FROM 
                ((tbl_MarkerData_Vegetation
                            INNER JOIN tbl_MarkerData
                                ON tbl_MarkerData_Vegetation.Point_ID = tbl_MarkerData.Point_ID)
                            INNER JOIN tbl_Events 
                                ON tbl_MarkerData.Event_ID = tbl_Events.Event_ID)
                            INNER JOIN geo_Locations 
                                ON tbl_Events.Location_ID = geo_Locations.GlobalID
            WHERE
                tbl_MarkerData_Vegetation.PercentCover IS NOT NULL
                AND tbl_Events.Event_Type = 'Marker Visit'
                AND Loc_Name_Short IN ('01_1', '01_2', '01_3', '01_4', '02_1', '02_2', '02_3', '02_4',
                '03_1', '03_2', '03_3', '03_4', '04_1', '04_2', '04_3', '04_4',
                '05_1', '05_2', '05_3', '05_4', '06_1', '06_2', '06_3', '06_4',
                '07_1', '07_2', '07_3', '07_4', '08_1', '08_2', '08_3', '08_4',
                '09_1', '09_2', '09_3', '09_4', '10_1', '10_2', '10_3', '10_4',
                '11_1', '11_2', '11_3', '11_4', '12_1', '12_2', '12_3', '12_4',
                '13_1', '13_2', '13_3', '13_4', '14_1', '14_2', '14_3', '14_4')
            ORDER BY Loc_Name_Short
            """
        
        outVal = connect_to_AccessDB(in_query, inDB)

        if outVal[0].lower() != "success function":
            print(f"WARNING - Function defineRecords_AbsCoverByPoint - Failed - Exiting Script- {timeFun()}")
            exit()

        outVal = outVal[1]
        print(f"Success: defineRecords_AbsCoverByPoint - {timeFun()}")

        # Calc Absolute Cover for each species and Location / Community
        cols = [
            "PercentCover",
            "MangroveSide_Cover_Overall",
            "MangroveSide_Cover_Tree",
            "MangroveSide_Cover_Shrub",
            "MangroveSide_Cover_Herb",
            "MarshSide_Cover_Overall",
            "MarshSide_Cover_Tree",
            "MarshSide_Cover_Shrub",
            "MarshSide_Cover_Herb",
        ]

        outVal[cols] = outVal[cols].apply(pd.to_numeric, errors="coerce")
        outVal["AbsolutePercCover"] = outVal.apply(compute_cover, axis=1)

        pivot_df = outVal.pivot_table(
            values='AbsolutePercCover',
            index=['CommunityType', 'VegetationType', 'SpeciesCode'],
            columns='Location_Name',
            aggfunc='sum'
        )

        # Append DataFrame to existing excel file
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_oldDB_{dateString}.xlsx")

        with pd.ExcelWriter(outFull, mode='a', engine="openpyxl") as writer:
            pivot_df.to_excel(writer, sheet_name=f"SOP8-2-AbsCov", index=True)

        scriptMsg = f"EXPORTED Table 8-2-AbsCov to {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)

        return "success function", outVal
    except:
        print(f"Error on defineRecords_AbsCoverByPoint - {timeFun()}")
        logging.exception("WARNING Script Failed - defineRecords_AbsCoverByPoint")
        return "Failed function - 'defineRecords_AbsCoverByPoint'"

#Calculate Absolute Cover By Stratum and Community type in table 'tbl_MarkerData'
def defineRecords_AbsCoverByStratum():
    try:
        in_query = """
        SELECT
            geo_Locations.Loc_Name AS Location_Name,
            tbl_Events.Event_ID,
            tbl_Events.Start_Date,

            tbl_MarkerData.MangroveSide_Cover_Overall,
            tbl_MarkerData.MangroveSide_Cover_Tree,
            tbl_MarkerData.MangroveSide_Cover_Shrub,
            tbl_MarkerData.MangroveSide_Cover_Herb,

            [MangroveSide_Cover_Overall] * ([MangroveSide_Cover_Tree] / 100)  AS AbsCover_Mangrove_Tree,
            [MangroveSide_Cover_Overall] * ([MangroveSide_Cover_Shrub] / 100) AS AbsCover_Mangrove_Shrub,
            [MangroveSide_Cover_Overall] * ([MangroveSide_Cover_Herb] / 100)  AS AbsCover_Mangrove_Herb,

            tbl_MarkerData.MarshSide_Cover_Overall,
            tbl_MarkerData.MarshSide_Cover_Tree,
            tbl_MarkerData.MarshSide_Cover_Shrub,
            tbl_MarkerData.MarshSide_Cover_Herb,

            [MarshSide_Cover_Overall] * ([MarshSide_Cover_Tree] / 100)  AS AbsCover_Marsh_Tree,
            [MarshSide_Cover_Overall] * ([MarshSide_Cover_Shrub] / 100) AS AbsCover_Marsh_Shrub,
            [MarshSide_Cover_Overall] * ([MarshSide_Cover_Herb] / 100)  AS AbsCover_Marsh_Herb

        FROM
            ((geo_Locations
                INNER JOIN tbl_Events 
                    ON tbl_Events.Location_ID = geo_Locations.GlobalID)
                INNER JOIN tbl_MarkerData 
                    ON tbl_MarkerData.Event_ID = tbl_Events.Event_ID)
                INNER JOIN tbl_MarkerData_Vegetation
                    ON tbl_MarkerData_Vegetation.Point_ID = tbl_MarkerData.Point_ID

        WHERE
            tbl_Events.Event_Type = 'Marker Visit'
            AND Loc_Name IN ('01_1', '01_2', '01_3', '01_4', '02_1', '02_2', '02_3', '02_4',
                '03_1', '03_2', '03_3', '03_4', '04_1', '04_2', '04_3', '04_4',
                '05_1', '05_2', '05_3', '05_4', '06_1', '06_2', '06_3', '06_4',
                '07_1', '07_2', '07_3', '07_4', '08_1', '08_2', '08_3', '08_4',
                '09_1', '09_2', '09_3', '09_4', '10_1', '10_2', '10_3', '10_4',
                '11_1', '11_2', '11_3', '11_4', '12_1', '12_2', '12_3', '12_4',
                '13_1', '13_2', '13_3', '13_4', '14_1', '14_2', '14_3', '14_4')

        ORDER BY
            geo_Locations.Loc_Name,
            tbl_Events.Start_Date;
        """

        outVal = connect_to_AccessDB(inQuery, inDB)

        if outVal[0].lower() != "success function":
            print("WARNING - Function defineRecords_AbsCoverByStratum - Failed - Exiting Script -  {timeFun()}")
            exit()
        
        outVal = outVal[1]
        # join with regions table

        # Append DataFrame to existing excel file
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_oldDB_{dateString}.xlsx")

        with pd.ExcelWriter(outFull, mode='a', engine="openpyxl") as writer:
            outVal.to_excel(writer, sheet_name="AbsCovByStratum", index=True)

        scriptMsg = f"EXPORTED Table AbsCovByStratum to {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        
        print(f"Success:  defineRecords_AbsCoverByStratum - {timeFun()}")
        return "success function", outVal

    except:
        print(f"Error on defineRecords_AbsCoverByStratum Function - {timeFun()}")
        logging.exception("WARNING Script Failed - defineRecords_AbsCoverByStratum")
        return "Failed function - 'defineRecords_AbsCoverByStratum'"

#Create  Figures - Absolute Cover By Region, By Community, By Strata
def figure_AbsCoverByStratum(inoutVal):
    try:

        #Open PoutVal to be copied to
        poutVal = PoutValPages(outPoutVal)

        # Process By Region
        regionlList = ['Turner River', 'Shark Slough', 'Taylor Slough']
        for count, region in enumerate(regionlList):

            #Subset By Region
            outValSub = inoutVal[inoutVal['Region'] == region]

            ###########################
            #Marsh DataFrame and Figure
            ###########################
            #Subset to marshoutVal fields
            marshoutVal = outValSub.loc[:,('Location_Name', 'MarshSide_Cover_Overall', 'AbsCover_Marsh_Tree', 'AbsCover_Marsh_Shrub','AbsCover_Marsh_Herb')]
            #Rename Fields
            marshoutVal.rename(columns={"AbsCover_Marsh_Tree": "Tree", "AbsCover_Marsh_Shrub": "Shrub", "AbsCover_Marsh_Herb": "Herb"}, inplace=True)

            #Set Index
            marshoutVal.set_index('Location_Name', inplace=True)

            #Subset to Mangrove fields
            mangroveoutVal = outValSub.loc[:, ('Location_Name', 'MangroveSide_Cover_Overall', 'AbsCover_Mangrove_Tree','AbsCover_Mangrove_Shrub', 'AbsCover_Mangrove_Herb')]
            # Rename Fields
            mangroveoutVal.rename(columns={"AbsCover_Mangrove_Tree": "Tree", "AbsCover_Mangrove_Shrub": "Shrub", "AbsCover_Mangrove_Herb": "Herb"}, inplace=True)

            #Set Index
            mangroveoutVal.set_index('Location_Name', inplace=True)

            ####################
            #Create the Figures:
            ####################
            plt.figure(figsize=(8, 6))
            ax1 = plt.subplot(2, 1, 1)
            marshoutVal.plot.bar(stacked=True, title="Marsh Side - " + region, xlabel="Marker Points", ylabel="Absolute Percent Cover (%)", color={'Shrub': 'olivedrab', 'Herb': 'gold', 'Tree': 'saddlebrown'}, ax=ax1)
            lgd = plt.legend(['Tree', 'Shrub', 'Herb'], loc='center left', bbox_to_anchor=(1, 0.5))
            plt.grid(axis='y')
            plt.ylim(0, 100)
            plt.tight_layout(pad=0.4)

            ax2 = plt.subplot(2, 1, 2)
            mangroveoutVal.plot.bar(stacked=True, title="Mangrove Side - " + region, xlabel="Marker Points", ylabel="Absolute Percent Cover (%)", color={'Shrub': 'olivedrab', 'Herb': 'gold', 'Tree': 'saddlebrown'}, ax=ax2)
            lgd = plt.legend(['Tree', 'Shrub', 'Herb'], loc='center left', bbox_to_anchor=(1, 0.5))
            plt.grid(axis='y')
            plt.ylim(0, 100)
            plt.tight_layout(pad=0.4)

            figure = mp.pyplot.gcf()
            poutVal.savefig(figure)

            scriptMsg = f"EXPORTED Figures Region: {region} - {timeFun()}"
            print(scriptMsg)
            logging.info(scriptMsg)

        poutVal.close()

        print(f"Success: figure_AbsCoverByStratum - {timeFun()}")
        return "success function"

    except:
        print(f"Error on figure_AbsCoverByStratum Function - {timeFun()}")
        logging.exception("WARNING Script Failed - figure_AbsCoverByStratum")
        return "Failed function - 'figure_AbsCoverByStratum'"

#Extract Mangrove Marsh Distance Records table 'tbl_MarkerData' where Event Type = 'Marker Visit'
def defineRecords_MarkerData():
    try:
        inQuery = "SELECT tbl_Event_Group.Event_Group_ID, tbl_Event_Group.Event_Group_Name, tbl_Event_Group.Start_Date, tbl_Event_Group.End_Date, tbl_Event_Group.Assessment, tbl_Events.Event_Type,"\
                " tbl_Events.Location_ID, geo_Locations.Region,geo_Locations.Segment, geo_Locations.Location_Name, tbl_MarkerData.Distance, tbl_MarkerData.Method"\
                " FROM geo_Locations INNER JOIN ((tbl_Event_Group INNER JOIN tbl_Events ON (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID) AND (tbl_Event_Group.Event_Group_ID = tbl_Events.Event_Group_ID))"\
                " INNER JOIN tbl_MarkerData ON (tbl_Events.Event_ID = tbl_MarkerData.Event_ID) AND (tbl_Events.Event_ID = tbl_MarkerData.Event_ID)) ON geo_Locations.GlobalID = tbl_Events.Location_ID"\
                " WHERE tbl_Events.Event_Type = 'Marker Visit' ORDER BY geo_Locations.Segment, geo_Locations.Location_Name, tbl_Events.Event_Type;"\

        outVal = connect_to_AccessDB(inQuery, inDB)
        if outVal[0].lower() != "success function":
            print(f"WARNING - Function defineRecords_MarkerData - Failed - Exiting Script - {timeFun()}")
            exit()
            
        outVal = outVal[1]
        print(f"Success:  defineRecords_MarkerData - {timeFun()}")
        return "success function", outVal

    except:
        print(f"Error on defineRecords_MarkderData Function - {timeFun()}")
        logging.exception("WARNING Script Failed - defineRecords_MarkderData")
        return "Failed function - 'defineRecords_MarkderData'"

# Summarize Mangrove Marsh Ecotone Values - Average Distance, Standard Error, Lower 95% Confidence Limit, Upper 95% Confidence Limit, Max and Min Values
## Calculate the Confidence Interval Upper 95% and Lower 95% using Students T Distribution
### Student T Distribution is defined as t_crit = np.abs(t.ppf((1-confidence)/2,dof))
def SummarizeFigure8_1(inoutVal):
    try:
        # Calculate Average, Standard Error, and count of Distance 
        outVal_8pt1 = (
            inoutVal.groupby(['Event_Group_ID', 'Region', 'Segment'], as_index=False)
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
        outFull = os.path.join(outputDir, f"MangroveMarsh_Export_oldDB_{dateString}.xlsx")
        outVal_8pt1.to_excel(outFull, sheet_name = 'SOP8-1', index=False)

        scriptMsg = f"EXPORTED Table 8-1 to: {outFull} - {timeFun()}"
        print(scriptMsg)
        logging.info(scriptMsg)
        return "success function", outVal_8pt1

    except:
        print(f"Error on SummarizeFigure8_1 Function - {timeFun()}")
        logging.exception("Error in SummarizeFigure8_1")
        return "Failed function - 'SummarizeFigure8_1'"

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