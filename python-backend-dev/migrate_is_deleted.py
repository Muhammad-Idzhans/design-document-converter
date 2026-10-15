"""
migrate_is_deleted.py

Database migration script to ensure all existing Cosmos DB task records 
have the `isDeleted: False` attribute so that:
- History page can filter active documents (`isDeleted == False`)
- Dashboard continues aggregating all historical tasks (active and inactive)
"""

import os
from dotenv import load_dotenv

load_dotenv()

try:
    from azure.cosmos import CosmosClient
except ImportError:
    print("Error: azure-cosmos is not installed. Please install it first.")
    exit(1)

COSMOS_URI = os.getenv("COSMOS_DB_URI")
COSMOS_KEY = os.getenv("COSMOS_DB_KEY")
DATABASE_NAME = os.getenv("COSMOS_DB_DATABASE", "design-doc-generator")
CONTAINER_NAME = os.getenv("COSMOS_DB_TASKS_CONTAINER", "tasks")

def run_migration():
    if not COSMOS_URI or not COSMOS_KEY:
        print("Error: COSMOS_DB_URI or COSMOS_DB_KEY is missing from .env.")
        return

    print(f"Connecting to Cosmos DB: {COSMOS_URI}")
    client = CosmosClient(COSMOS_URI, credential=COSMOS_KEY)
    database = client.get_database_client(DATABASE_NAME)
    container = database.get_container_client(CONTAINER_NAME)

    print(f"Connected to container '{CONTAINER_NAME}' in database '{DATABASE_NAME}'.")
    print("Scanning task documents...")

    query = "SELECT * FROM c"
    items = list(container.query_items(query=query, enable_cross_partition_query=True))

    total = len(items)
    updated_count = 0
    already_flagged = 0

    print(f"Found {total} total task records.")

    for item in items:
        task_id = item.get("taskId") or item.get("id")
        # Check if isDeleted is missing
        if "isDeleted" not in item or item.get("isDeleted") is None:
            item["isDeleted"] = False
            container.upsert_item(item)
            updated_count += 1
            print(f"  [UPDATED] Task {task_id}: Set isDeleted = False")
        else:
            already_flagged += 1

    print("\n" + "=" * 50)
    print("Migration Summary:")
    print(f"  Total tasks scanned:    {total}")
    print(f"  Tasks updated:          {updated_count}")
    print(f"  Already had isDeleted:  {already_flagged}")
    print("=" * 50)

if __name__ == "__main__":
    run_migration()
