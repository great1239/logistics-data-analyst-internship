"""
generate_dataset.py
Generates a realistic, enterprise-grade logistics shipment dataset for the
Logistics Data Analyst Internship project.
"""

import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import random
import os

def generate_logistics_data(num_records=5000, random_state=42):
    np.random.seed(random_state)
    random.seed(random_state)

    cities = ['Delhi', 'Mumbai', 'Bengaluru', 'Chennai', 'Kolkata', 'Hyderabad', 'Pune', 'Ahmedabad']
    carriers = ['BlueDart Logistics', 'Delhivery Express', 'SafeXpress', 'TCI Freight', 'Gati KWE']
    shipping_modes = ['Express', 'Standard', 'Economy']
    vehicle_types = ['Electric Van', 'LCV (Light Commercial)', 'MCV (Medium Commercial)', 'Heavy Truck (16T)']
    weather_types = ['Clear', 'Rain', 'Fog', 'Storm']
    traffic_levels = ['Low', 'Medium', 'High', 'Severe']

    # City pair distance matrix (approximate km)
    city_coords = {
        'Delhi': (28.6139, 77.2090),
        'Mumbai': (19.0760, 72.8777),
        'Bengaluru': (12.9716, 77.5946),
        'Chennai': (13.0827, 80.2707),
        'Kolkata': (22.5726, 88.3639),
        'Hyderabad': (17.3850, 78.4867),
        'Pune': (18.5204, 73.8567),
        'Ahmedabad': (23.0225, 72.5714)
    }

    def calc_dist(c1, c2):
        lat1, lon1 = city_coords[c1]
        lat2, lon2 = city_coords[c2]
        # Approximate euclidean distance on grid * 111 km/deg * road tortuosity factor 1.25
        d = np.sqrt((lat1 - lat2)**2 + (lon1 - lon2)**2) * 111 * 1.25
        return round(float(d), 1)

    start_date = datetime(2026, 1, 1, 8, 0, 0)

    records = []
    for i in range(num_records):
        shipment_id = f"SHP-{100000 + i}"
        origin = random.choice(cities)
        dest = random.choice([c for c in cities if c != origin])
        distance_km = calc_dist(origin, dest)

        mode = np.random.choice(shipping_modes, p=[0.25, 0.50, 0.25])
        carrier = random.choice(carriers)

        # Vehicle type depends partly on shipping mode
        if mode == 'Express' and distance_km < 400:
            veh = np.random.choice(['Electric Van', 'LCV (Light Commercial)'], p=[0.4, 0.6])
        elif distance_km > 1000:
            veh = np.random.choice(['Heavy Truck (16T)', 'MCV (Medium Commercial)'], p=[0.7, 0.3])
        else:
            veh = random.choice(vehicle_types)

        # Shipment weight and volume
        if 'Van' in veh:
            weight = np.random.uniform(50, 1000)
            volume = weight / np.random.uniform(150, 250)
            avg_speed = np.random.uniform(55, 75)
        elif 'LCV' in veh:
            weight = np.random.uniform(500, 3500)
            volume = weight / np.random.uniform(200, 300)
            avg_speed = np.random.uniform(45, 65)
        elif 'MCV' in veh:
            weight = np.random.uniform(3000, 9000)
            volume = weight / np.random.uniform(250, 350)
            avg_speed = np.random.uniform(40, 55)
        else:
            weight = np.random.uniform(8000, 18000)
            volume = weight / np.random.uniform(300, 400)
            avg_speed = np.random.uniform(35, 50)

        dispatch_time = start_date + timedelta(
            days=random.randint(0, 180),
            hours=random.randint(0, 23),
            minutes=random.randint(0, 59)
        )

        warehouse_dock_hours = round(np.random.gamma(shape=2.5, scale=1.5), 1) + 0.5
        planned_transit_hours = round((distance_km / avg_speed) + warehouse_dock_hours + (2 if mode == 'Standard' else 0), 1)

        weather = np.random.choice(weather_types, p=[0.65, 0.20, 0.10, 0.05])
        traffic = np.random.choice(traffic_levels, p=[0.30, 0.40, 0.20, 0.10])

        weather_delay = {'Clear': 0, 'Rain': np.random.uniform(1, 4), 'Fog': np.random.uniform(2, 6), 'Storm': np.random.uniform(5, 14)}[weather]
        traffic_delay = {'Low': 0, 'Medium': np.random.uniform(0.5, 2), 'High': np.random.uniform(2, 5), 'Severe': np.random.uniform(4, 9)}[traffic]
        carrier_reliability = {'BlueDart Logistics': 0.90, 'Delhivery Express': 0.86, 'SafeXpress': 0.82, 'TCI Freight': 0.78, 'Gati KWE': 0.75}[carrier]

        mechanical_delay = np.random.exponential(scale=3.0) if np.random.rand() > carrier_reliability else 0
        warehouse_congestion_delay = max(0, warehouse_dock_hours - 4) * 0.75

        total_delay = round(weather_delay + traffic_delay + mechanical_delay + warehouse_congestion_delay + np.random.normal(0, 0.8), 2)
        total_delay = max(-2.0, total_delay) # early delivery possible up to 2 hrs
        actual_transit_hours = max(1.0, round(planned_transit_hours + total_delay, 1))

        # Shipping cost formula (base + per_km + per_kg + fuel)
        base_fare = 500
        km_rate = 1.8 if 'Heavy' in veh else (1.4 if 'MCV' in veh else 1.1)
        kg_rate = 0.45 if mode == 'Express' else 0.25
        fuel_price = round(np.random.uniform(92.0, 103.5), 2)
        fuel_liters = round((distance_km / (3.5 if 'Heavy' in veh else 8.0)) * np.random.uniform(0.9, 1.2), 1)
        shipping_cost = round(base_fare + (distance_km * km_rate) + (weight * kg_rate) + (fuel_liters * (fuel_price / 15)), 2)

        is_delayed = 1 if total_delay > 0.5 else 0

        # Inject realistic dirty data
        weight_raw = weight
        dock_hours_raw = warehouse_dock_hours
        if random.random() < 0.035: # 3.5% missing weights
            weight_raw = np.nan
        if random.random() < 0.04:  # 4% missing dock hours
            dock_hours_raw = np.nan
        if random.random() < 0.01:  # 1% extreme outlier sensor glitches
            actual_transit_hours = actual_transit_hours * 10

        carrier_entry = carrier
        if random.random() < 0.03:  # casing / space inconsistency
            carrier_entry = carrier.lower() + " "

        records.append({
            'Shipment_ID': shipment_id,
            'Origin_City': origin,
            'Destination_City': dest,
            'Distance_km': distance_km,
            'Shipping_Mode': mode,
            'Carrier_Name': carrier_entry,
            'Vehicle_Type': veh,
            'Shipment_Weight_kg': round(weight_raw, 2) if pd.notnull(weight_raw) else np.nan,
            'Package_Volume_m3': round(volume, 3),
            'Dispatch_Timestamp': dispatch_time.strftime('%Y-%m-%d %H:%M:%S'),
            'Warehouse_Dock_Hours': round(dock_hours_raw, 2) if pd.notnull(dock_hours_raw) else np.nan,
            'Weather_Condition': weather,
            'Traffic_Level': traffic,
            'Fuel_Consumed_Liters': fuel_liters,
            'Fuel_Price_INR': fuel_price,
            'Planned_Transit_Hours': planned_transit_hours,
            'Actual_Transit_Hours': actual_transit_hours,
            'Delay_Hours': total_delay,
            'Shipping_Cost_INR': shipping_cost,
            'Is_Delayed': is_delayed
        })

    df = pd.DataFrame(records)
    return df

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    os.makedirs(data_dir, exist_ok=True)
    raw_path = os.path.join(data_dir, 'raw_logistics_data.csv')

    df = generate_logistics_data(5000)
    df.to_csv(raw_path, index=False)
    print(f"Generated raw logistics dataset: {raw_path} with {len(df)} rows and {len(df.columns)} columns.")
