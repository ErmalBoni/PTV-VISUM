# Script Created by Ermal Sylejmani, ChatGPT, ClaudeAI - 28/08/2024
# Name: CreateShapefilesFromVisum.py
# Description: This Python script connects to the PTV Visum software and automates the export of traffic model data into CSV files and Shapefiles.
# It retrieves and exports data for nodes, links, zones, and bus stops, including a range of useful attributes.
# The script handles errors and logs progress to ensure smooth operation.

# Key Features:
# - Connects to PTV Visum and loads a specified traffic model.
# - Exports data into CSV files and Shapefiles for nodes, links, zones, and bus stops.
# - Attributes exported include node numbers, link capacities, zone coordinates, and bus stop details.
# - Comprehensive logging and error handling for robust performance.

# Instructions:
# 1. Ensure you have the required Python packages installed:
#    pip install pywin32 pandas geopandas shapely chardet
# 2. Update the `model_path` variable with the full path to your PTV Visum model file (line 98).
#    # Update: Add your model path here
#    model_path = r"C:\path\to\your\model.ver"
# 3. Run the script using a Python interpreter or via PowerShell. To run via PowerShell, use the command:
#    python .\CreateShapefilesFromVisum.py
# 4. The following CSV and Shapefiles will be created in the script's directory:
#    - `Nodes.csv` and `Nodes.shp`: Contains details about nodes including node number, control type, type number, and coordinates.
#    - `Links.csv` and `Links.shp`: Includes information on links such as link number, from/to node numbers, length, capacity, free flow speed, and volume.
#    - `Zones.csv` and `Zones.shp`: Lists zones with zone number and coordinates.
#    - `StopPoints.csv` and `StopPoints.shp`: Provides data on stop points including stop point number, coordinates, name, associated node number, number of lines, and TSysSet.

# 5. For different projects or changes in attributes, update the following lines:
#    - Line 57: Modify the attributes retrieved for nodes if you need different node details.
#    - Line 73: Adjust the attributes for links to match the required data for your project.
#    - Line 89: Change the attributes for zones if additional or different zone data is needed.
#    - Line 105: Update the attributes for stop points based on your specific requirements.

import win32com.client
import csv
import os
import logging
import time
import shutil
from time import time as timer
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point, LineString

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
logger.propagate = False

CONTROL_TYPE_MAP = {
    0: "unknown", 1: "Uncontrolled", 2: "Two-way stop", 3: "Signalized",
    4: "All-way stop", 5: "Roundabout", 6: "Two-way yield"
}

def clear_com_cache():
    try:
        gen_py_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", "gen_py")
        if os.path.exists(gen_py_dir):
            shutil.rmtree(gen_py_dir)
            logger.info("Cleared win32com cache.")
        else:
            logger.info("win32com cache directory not found.")
    except Exception as e:
        logger.error(f"Failed to clear win32com cache: {e}")

def Init():
    try:
        start_time = timer()
        Visum = win32com.client.Dispatch('Visum.Visum')
        end_time = timer()
        logger.info(f"Visum application launched successfully in {end_time - start_time:.2f} seconds.")
        return Visum
    except Exception as e:
        logger.error(f"Error connecting to Visum: {e}")
        logger.info("Attempting to clear COM cache and retry...")
        clear_com_cache()
        try:
            time.sleep(2)
            start_time = timer()
            Visum = win32com.client.Dispatch('Visum.Visum')
            end_time = timer()
            logger.info(f"Visum connected successfully after clearing COM cache in {end_time - start_time:.2f} seconds.")
            return Visum
        except Exception as e:
            logger.error(f"Error connecting to Visum after clearing COM cache: {e}")
            return None

def load_model(Visum, model_path):
    try:
        start_time = timer()
        Visum.LoadVersion(model_path)
        end_time = timer()
        logger.info(f"Model loaded successfully from {model_path} in {end_time - start_time:.2f} seconds.")
        return True
    except Exception as e:
        logger.error(f"Failed to load model from {model_path}: {e}")
        return False

def export_to_csv(data, headers, filename):
    try:
        start_time = timer()
        current_dir = os.path.dirname(os.path.abspath(__file__))
        output_csv_path = os.path.join(current_dir, filename)
        with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow(headers)
            for row in data:
                new_row = [f"{int(item)}" if isinstance(item, float) and item.is_integer() 
                           else f"{item:.6f}" if isinstance(item, float) 
                           else str(item) for item in row]
                writer.writerow(new_row)
        end_time = timer()
        logger.info(f"CSV file created successfully: {output_csv_path} in {end_time - start_time:.2f} seconds.")
        return output_csv_path
    except Exception as e:
        logger.error(f"Error creating CSV file {filename}: {e}")
        return None

def create_shapefile_from_csv(csv_file, shapefile_name, geom_type, x_field=None, y_field=None, geom_fields=None):
    try:
        df = pd.read_csv(csv_file, delimiter=';', encoding='utf-8')
        
        if geom_type == "POINT":
            df['geometry'] = df.apply(lambda row: Point(float(row[x_field]), float(row[y_field])), axis=1)
        elif geom_type == "LINE" and geom_fields:
            df['geometry'] = df.apply(lambda row: LineString([
                (float(row[geom_fields[0]]), float(row[geom_fields[1]])),
                (float(row[geom_fields[2]]), float(row[geom_fields[3]]))
            ]), axis=1)
        
        gdf = gpd.GeoDataFrame(df, geometry='geometry', crs="EPSG:4326")
        
        shapefile_dir = os.path.join(os.path.dirname(csv_file), "Shapefiles")
        if not os.path.exists(shapefile_dir):
            os.makedirs(shapefile_dir)
        
        shapefile_path = os.path.join(shapefile_dir, f"{shapefile_name}.shp")
        gdf.to_file(shapefile_path)
        
        logger.info(f"Shapefile {shapefile_name}.shp created successfully in Shapefiles directory.")
    except Exception as e:
        logger.error(f"Error creating shapefile {shapefile_name}: {e}")

def export_nodes_with_control_type(Visum):
    try:
        logger.info("Exporting nodes...")
        headers = ["NodeNo", "CtrlType", "TypeNo", "XCoord", "YCoord"]
        start_time = timer()
        Nodes = Visum.Net.Nodes.GetMultipleAttributes(["No", "ControlType", "TypeNo", "XCoord", "YCoord"])
        end_time = timer()
        logger.debug(f"Retrieved {len(Nodes)} nodes in {end_time - start_time:.2f} seconds.")
        
        if not Nodes:
            logger.warning("No nodes were retrieved from Visum.")
            return None

        processed_nodes = [[node[0], CONTROL_TYPE_MAP.get(node[1], "Unknown"), node[2], node[3], node[4]] for node in Nodes]
        csv_file = export_to_csv(processed_nodes, headers, "Nodes.csv")
        
        if csv_file:
            create_shapefile_from_csv(csv_file, "Nodes", "POINT", x_field="XCoord", y_field="YCoord")
        
        return csv_file
    except Exception as e:
        logger.error(f"Error exporting nodes: {e}")
        return None

def export_links(Visum):
    try:
        logger.info("Exporting links...")
        headers = ["LinkID", "No", "FromNode", "ToNode", "Length", "Capacity", "FreeFlowSpd", "NumLanes", "VolumePrT", "FromX", "FromY", "ToX", "ToY"]
        available_attributes = ["No", "FromNodeNo", "ToNodeNo", "Length", "CapPrT", "V0PrT", "NumLanes", "VolVehPrT(AP)"]  # Adjusted VolVehPrT with AP subattribute
        
        # Adding XY coordinates for FromNode and ToNode
        additional_attributes = ["FromNode\\XCoord", "FromNode\\YCoord", "ToNode\\XCoord", "ToNode\\YCoord"]

        logger.debug(f"Attempting to retrieve attributes: {available_attributes + additional_attributes}")
        start_time = timer()
        Links = Visum.Net.Links.GetMultipleAttributes(available_attributes + additional_attributes)
        end_time = timer()
        logger.debug(f"Retrieved {len(Links)} links in {end_time - start_time:.2f} seconds.")
        
        if not Links or len(Links) == 0:
            logger.warning("No links were retrieved from Visum.")
            return None
        
        processed_links = []
        for link in Links:
            if len(link) != len(available_attributes + additional_attributes):
                logger.warning(f"Link data mismatch. Expected {len(available_attributes + additional_attributes)} attributes, got {len(link)}")
                continue
            
            # Create a unique Link ID from FromNode and ToNode
            link_id = f"{int(link[1])}_{int(link[2])}"
            processed_link = [link_id] + list(link)
            processed_links.append(processed_link)
        
        logger.debug(f"Processed {len(processed_links)} links for export.")
        
        # Export to CSV
        csv_file = export_to_csv(processed_links, headers, "Links.csv")
        if csv_file:
            logger.info(f"Links CSV successfully created: {csv_file}")
            
            # Now create the shapefile using the new Link ID and coordinates
            create_shapefile_from_csv(csv_file, "Links", "LINE", geom_fields=["FromX", "FromY", "ToX", "ToY"])
        else:
            logger.error("Failed to create Links CSV.")
        
        return csv_file
    except AttributeError as e:
        logger.error(f"Visum attribute error: {e}. Check if all attributes are valid for your Visum version.")
    except win32com.client.pywintypes.com_error as e:
        logger.error(f"COM error when exporting links: {e}")
    except Exception as e:
        logger.error(f"Unexpected error exporting links: {e}")
    return None

def export_zones(Visum):
    try:
        logger.info("Exporting zones...")
        headers = ["ZoneNo", "XCoord", "YCoord"]
        start_time = timer()
        Zones = Visum.Net.Zones.GetMultipleAttributes(["No", "XCoord", "YCoord"])
        end_time = timer()
        logger.debug(f"Retrieved {len(Zones)} zones in {end_time - start_time:.2f} seconds.")
        
        if not Zones:
            logger.warning("No zones were retrieved from Visum.")
            return None
        
        csv_file = export_to_csv(Zones, headers, "Zones.csv")
        
        if csv_file:
            create_shapefile_from_csv(csv_file, "Zones", "POINT", x_field="XCoord", y_field="YCoord")
        
        return csv_file
    except Exception as e:
        logger.error(f"Error exporting zones: {e}")
        return None

def export_stop_points(Visum):
    try:
        logger.info("Exporting stop points...")
        headers = ["StopPtNo", "XCoord", "YCoord", "Name", "NodeNo", "NumLines", "TSysSet"]
        start_time = timer()
        StopPoints = Visum.Net.StopPoints.GetMultipleAttributes(["No", "XCoord", "YCoord", "Name", "NodeNo", "NumLines", "TSysSet"])
        end_time = timer()
        logger.debug(f"Retrieved {len(StopPoints)} stop points in {end_time - start_time:.2f} seconds.")
        
        if not StopPoints:
            logger.warning("No stop points were retrieved from Visum.")
            return None
        
        csv_file = export_to_csv(StopPoints, headers, "StopPoints.csv")
        
        if csv_file:
            create_shapefile_from_csv(csv_file, "StopPoints", "POINT", x_field="XCoord", y_field="YCoord")
        
        return csv_file
    except Exception as e:
        logger.error(f"Error exporting stop points: {e}")
        return None

def main():
    # Update: Add your model path here
    model_path = r"C:\path\to\your\model.ver"  
    Visum = Init()
    if Visum and load_model(Visum, model_path):
        try:
            nodes_csv = export_nodes_with_control_type(Visum)
            if nodes_csv:
                time.sleep(5)
                logger.info("Nodes exported. Pausing for 5 seconds...")

            links_csv = export_links(Visum)
            if links_csv:
                time.sleep(5)
                logger.info("Links exported. Pausing for 5 seconds...")

            zones_csv = export_zones(Visum)
            if zones_csv:
                time.sleep(5)
                logger.info("Zones exported. Pausing for 5 seconds...")

            stop_points_csv = export_stop_points(Visum)
            if stop_points_csv:
                time.sleep(5)
                logger.info("Stop points exported. Pausing for 5 seconds...")

            logger.info("Export process completed.")
        
        except Exception as e:
            logger.error(f"Error in the export process: {e}")
    else:
        logger.error("Failed to initialize Visum or load the model.")

if __name__ == "__main__":
    main()
