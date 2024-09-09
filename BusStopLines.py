# Script Created by Ermal Sylejmani, ChatGPT - 09/09/2024
# Name: BusStopLines.py
# Description: This Python script connects to PTV Visum and automates the creation of a bus line with forward and backward routes.
# It reads stop IDs from a CSV file and generates a bus line based on the shortest path between stops, using distance as the default criterion.
# Additionally, time profiles for each route are created, and error handling and logging are provided for smooth execution.

# Key Features:
# - Connects to PTV Visum and loads a specified transport model.
# - Creates bus routes (forward and backward) based on the shortest path in distance between stops.
# - The shortest path can be switched to time-based instead of distance-based if necessary.
# - Generates time profiles for both forward and backward routes.
# - Supports multiple transport modes (e.g., bus, tram, metro).
# - Comprehensive logging and error handling for robust performance.

# Instructions:
# 1. Ensure you have the required Python packages installed:
#    pip install pywin32 pandas
# 2. Update the `model_path` variable with the full path to your PTV Visum model file.
#    Example:
#    model_path = r"C:\path\to\your\model.ver"
# 3. Update the `csv_path` variable with the full path to the CSV file containing stop IDs.
#    Example:
#    csv_path = r"C:\path\to\your\stops.csv"
# 4. Ensure the CSV file is structured as follows:
#    - Row 1: Header (e.g., Stop_1, Stop_2)
#    - Row 2: Stop IDs for the forward route.
#    - Row 3: Stop IDs for the backward route.
#
#    Example CSV structure:
#    ```
#    Stop_1,Stop_2
#    Stop1,Stop2, Stop3,....,n   # Forward direction
#    Stop1,Stop2, Stop3,....,n   # Backward direction
#    ```
# 5. By default, the script uses the shortest path based on distance between stops.
#    If you want to switch to using the shortest path based on time, change the following line:
#    - Line 80: Replace `C.ShortestPathCriterion_LinkLength` with `C.ShortestPathCriterion_TravelTime`.

# 6. The script can be used for other transport modes as well (e.g., tram, metro). Simply update the `tsys_code` variable (line 105).
#    Example:
#    tsys_code = "TRAM"  # or "METRO", "TRAIN", etc.

import win32com.client
import logging
import csv
from time import time as timer
import os

# Initialize logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

C = win32com.client.constants

def init_visum():
    """Initialize connection to Visum."""
    try:
        start_time = timer()
        Visum = win32com.client.Dispatch('Visum.Visum')
        end_time = timer()
        logger.info(f"Visum started successfully in {end_time - start_time:.2f} seconds.")
        return Visum
    except Exception as e:
        logger.error(f"Error starting Visum: {e}")
        return None

def load_model(Visum, model_path):
    """Load the transport model in Visum."""
    if not os.path.exists(model_path):
        logger.error(f"Model path {model_path} does not exist.")
        return False
    try:
        start_time = timer()
        Visum.LoadVersion(model_path)
        end_time = timer()
        logger.info(f"Model successfully loaded from {model_path} in {end_time - start_time:.2f} seconds.")
        return True
    except Exception as e:
        logger.error(f"Error loading the model from {model_path}: {e}")
        return False

def load_stops_from_csv(csv_file):
    """Load stop IDs from a CSV file."""
    try:
        with open(csv_file, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # Skip header
            rows = list(reader)

            stop_ids_forward = [int(stop_id) for stop_id in rows[0] if stop_id.strip()]
            stop_ids_backward = [int(stop_id) for stop_id in rows[1] if stop_id.strip()]

            logger.info(f"Forward Stops: {stop_ids_forward}")
            logger.info(f"Backward Stops: {stop_ids_backward}")
            return stop_ids_forward, stop_ids_backward
    except Exception as e:
        logger.error(f"Error loading stops from CSV: {e}")
        return [], []

def create_line_and_routes(Visum, line_name, tsys_code, stop_ids_forward, stop_ids_backward):
    """Create a new line with forward and backward routes in Visum."""
    try:
        logger.info(f"Creating new line '{line_name}'.")

        Net = Visum.Net
        line = Net.AddLine(line_name, tsys_code)
        
        direction_forward = Net.Directions.ItemByKey(">")
        direction_backward = Net.Directions.ItemByKey("<")

        # Create forward route
        route_forward = Visum.CreateNetElements()
        for stop_id in stop_ids_forward:
            stop_point = Net.StopPoints.ItemByKey(stop_id)
            if stop_point:
                route_forward.Add(stop_point)
            else:
                logger.error(f"Stop point with ID {stop_id} not found for forward route.")

        # Create backward route
        route_backward = Visum.CreateNetElements()
        for stop_id in stop_ids_backward:
            stop_point = Net.StopPoints.ItemByKey(stop_id)
            if stop_point:
                route_backward.Add(stop_point)
            else:
                logger.error(f"Stop point with ID {stop_id} not found for backward route.")

        routesearchparameters = Visum.IO.CreateNetReadRouteSearchTSys()
        routesearchparameters.SetAttValue("HowToHandleIncompleteRoute", C.RouteSearchHandleIncompleteRouteTSearchShortestPath)
        
        # Default to shortest path by distance, can switch to time-based
        routesearchparameters.SetAttValue("ShortestPathCriterion", C.ShortestPathCriterion_LinkLength)  # Change to C.ShortestPathCriterion_TravelTime if needed
        routesearchparameters.SetAttValue("IncludeBlockedLinks", False)
        routesearchparameters.SetAttValue("IncludeBlockedTurns", False)

        lineroute_forward = Net.AddLineRoute(f"{line_name}_forward", line, direction_forward, route_forward, routesearchparameters)
        lineroute_backward = Net.AddLineRoute(f"{line_name}_backward", line, direction_backward, route_backward, routesearchparameters)

        logger.info(f"Routes for '{line_name}' successfully created.")

        create_time_profile(Visum, lineroute_forward)
        create_time_profile(Visum, lineroute_backward)

        return line
    except Exception as e:
        logger.error(f"Error creating line and routes for '{line_name}': {e}")
        return None

def create_time_profile(Visum, line_route):
    """Create a time profile for the line route."""
    try:
        logger.info(f"Creating time profile for line route '{line_route.AttValue('Name')}'.")
        time_profile = Visum.Net.AddTimeProfile(f"TimeProfile_{line_route.AttValue('Name')}", line_route)

        for item in time_profile.TimeProfileItems.GetAll():
            item.SetAttValue("Arr", 120)
            item.SetAttValue("Dep", 120)

        logger.info(f"Time profile for line route '{line_route.AttValue('Name')}' created successfully.")
        return time_profile
    except Exception as e:
        logger.error(f"Error creating time profile: {e}")
        return None

def main():
    model_path = r"C:\path\to\your\model.ver" #Update path
    csv_path = r"C:\path\to\your\stops.csv"   #Update path
    
    Visum = init_visum()
    if Visum and load_model(Visum, model_path):
        try:
            line_name = "BusLine1"
            tsys_code = "BUS"  # Can also be "TRAM", "METRO", "TRAIN", etc.

            stop_ids_forward, stop_ids_backward = load_stops_from_csv(csv_path)

            if not stop_ids_forward or not stop_ids_backward:
                logger.error("No stops loaded from CSV. Exiting.")
                return

            new_line = create_line_and_routes(Visum, line_name, tsys_code, stop_ids_forward, stop_ids_backward)

            if new_line:
                Visum.SaveVersion(model_path)
                logger.info("Model saved successfully.")
            else:
                logger.error("Failed to create line. Model not saved.")
        except Exception as e:
            logger.error(f"Error during line creation process: {e}")
        finally:
            Visum = None
    else:
        logger.error("Could not initialize Visum or load model.")

if __name__ == "__main__":
    main()
