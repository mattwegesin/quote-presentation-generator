import os
import json
import time
import requests
import logging

logger = logging.getLogger(__name__)

def process_monday_gsa_duplicate_async(new_item_id, inn_code, monday_token):
    """Background task to clean up subitems and update status on a newly duplicated GSA Monday item."""
    board_id = "18424851511"
    monday_headers = {
        "Authorization": monday_token, 
        "API-Version": "2024-01", 
        "Content-Type": "application/json"
    }

    try:
        # Clean up subitems from the duplicated item FIRST (before status changes trigger automations)
        subitems = []
        max_retries = 12
        
        for attempt in range(max_retries):
            sub_query = """
            query ($itemId: [ID!]) {
              items (ids: $itemId) {
                subitems { id }
              }
            }
            """
            sub_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                "query": sub_query,
                "variables": {"itemId": [new_item_id]}
            })
            try:
                subitems = sub_resp.json().get('data', {}).get('items', [{}])[0].get('subitems', [])
                if subitems:
                    break
            except Exception:
                pass
            time.sleep(2.5)
            
        if subitems:
            logger.info(f"Background Thread: Found {len(subitems)} subitems on duplicated GSA item {new_item_id}. Deleting them...")
            del_mutations = []
            for idx, sub in enumerate(subitems):
                del_mutations.append(f"del{idx}: delete_item (item_id: {sub['id']}) {{ id }}")
            
            chunk_size = 30
            for i in range(0, len(del_mutations), chunk_size):
                chunk = del_mutations[i:i + chunk_size]
                del_query = "mutation { " + " ".join(chunk) + " }"
                requests.post("https://api.monday.com/v2", headers=monday_headers, json={"query": del_query})
        else:
            logger.info(f"Background Thread: No subitems found on duplicated GSA item {new_item_id} after waiting.")
            
        # Update column values and move group (Status: Contract Sent)
        update_query = """
        mutation ($itemId: ID!, $boardId: ID!, $group: String!, $colVals: JSON!) {
          move_item_to_group (item_id: $itemId, group_id: $group) { id }
          change_multiple_column_values (item_id: $itemId, board_id: $boardId, column_values: $colVals) { id }
        }
        """
        col_vals = json.dumps({
            "name": inn_code,
            "color_mm5zq7cg": {"label": "GSA ONLY"},
            "color_mm5w8ehb": {"label": "Gathering Info"},
            "status": {"label": "Sales"},
            "color_mm65xpj1": {"label": "Active"},
            "color_mm6sy6mm": {"label": "Contract Sent"}
        })
        
        update_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
            "query": update_query,
            "variables": {
                "itemId": new_item_id,
                "boardId": board_id,
                "group": "topics",
                "colVals": col_vals
            }
        })
        up_data = update_resp.json()
        if "errors" in up_data:
            logger.error(f"Monday Update Duplicated GSA Item Error: {up_data['errors']}")
        else:
            logger.info(f"Successfully moved and updated GSA item {new_item_id} (Status: Contract Sent).")
    except Exception as e:
        logger.error(f"Failed in process_monday_gsa_duplicate_async: {e}")
