import win32com.client
import pandas as pd
import time
from datetime import datetime, timezone
import os
SERVER_NAME = "Matrikon.OPC.Simulation.1"

ITEM_LIST = [
    "Random.Real8",
    "Random.Int4",
    "Bucket Brigade.Real8",
    "Saw-toothed Waves.Real8",
    "Random.Real8",
    "Random.Int4",
    "Bucket Brigade.Real8",
    "Saw-toothed Waves.Real8",
    "Random.Real8",
    "Random.Int4"
]

INTERVAL_SECONDS = 60
OUTPUT_DIR = "matrikon_logs"


os.makedirs(OUTPUT_DIR, exist_ok=True)


def get_hourly_filename():
    """Generate hourly CSV filename"""
    current_time = datetime.now()
    return os.path.join(
        OUTPUT_DIR,
        f"Matrikon_OPC_Log_{current_time.strftime('%Y-%m-%d_%H')}.csv"
    )


def initialize_csv(path):
    """Create CSV with header if file does not exist"""
    if not os.path.isfile(path):
        columns = [
            "Local_Timestamp",
            "UTC_Epoch_Time",
            "Tag_01", "Tag_02", "Tag_03", "Tag_04", "Tag_05",
            "Tag_06", "Tag_07", "Tag_08", "Tag_09", "Tag_10"
        ]
        pd.DataFrame(columns=columns).to_csv(path, index=False)


def connect_opc_server():
    """Connect to Matrikon OPC DA server"""
    opc_client = win32com.client.Dispatch("OPC.Automation.1")
    opc_client.Connect(SERVER_NAME)
    return opc_client


def setup_group(opc_client):
    """Create OPC group and add items"""
    group = opc_client.OPCGroups.Add("LoggerGroup")
    group.UpdateRate = 1000
    group.IsActive = True

    opc_items = group.OPCItems
    for index, tag in enumerate(ITEM_LIST, start=1):
        opc_items.AddItem(tag, index)

    return opc_items


def read_tag_values(opc_items):
    """Read values from OPC items"""
    values = []
    for index in range(1, len(ITEM_LIST) + 1):
        values.append(opc_items.Item(index).Value)
    return values


def main():
    print("Starting Matrikon OPC CSV Logger")

    opc = connect_opc_server()
    print("OPC Server Connected")

    opc_items = setup_group(opc)
    print("Tags registered. Logging started.")

    active_file = None

    while True:
        csv_path = get_hourly_filename()

        if csv_path != active_file:
            initialize_csv(csv_path)
            active_file = csv_path
            print(f"New hourly file created: {csv_path}")

        local_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        utc_epoch = int(datetime.now(timezone.utc).timestamp())

        tag_data = read_tag_values(opc_items)
        record = [local_time, utc_epoch] + tag_data

        pd.DataFrame([record]).to_csv(
            csv_path,
            mode="a",
            index=False,
            header=False
        )

        print(f"Data logged at {local_time}")
        time.sleep(INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
