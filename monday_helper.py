import os
import json
import time
import requests
import logging

logger = logging.getLogger(__name__)

def process_monday_gsa_duplicate_async(new_item_id, inn_code, monday_token, gsa_quote_id=None, total_monthly=None):
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
        max_retries = 15
        last_count = -1
        stable_polls = 0
        
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
                current_subitems = sub_resp.json().get('data', {}).get('items', [{}])[0].get('subitems', [])
                current_count = len(current_subitems)
                
                if current_count > 0:
                    if current_count == last_count:
                        stable_polls += 1
                        if stable_polls >= 2:
                            subitems = current_subitems
                            break
                    else:
                        stable_polls = 0
                last_count = current_count
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
                
            # Verify deletion and delete any stragglers
            for _ in range(5):
                time.sleep(2.0)
                check_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                    "query": sub_query,
                    "variables": {"itemId": [new_item_id]}
                })
                try:
                    remaining = check_resp.json().get('data', {}).get('items', [{}])[0].get('subitems', [])
                    if not remaining:
                        logger.info("Background Thread: Verified all subitems deleted.")
                        break
                    else:
                        logger.warning(f"Background Thread: {len(remaining)} subitems still remain. Retrying delete...")
                        del_mutations = []
                        for idx, sub in enumerate(remaining):
                            del_mutations.append(f"del{idx}: delete_item (item_id: {sub['id']}) {{ id }}")
                        
                        for i in range(0, len(del_mutations), chunk_size):
                            chunk = del_mutations[i:i + chunk_size]
                            del_query = "mutation { " + " ".join(chunk) + " }"
                            requests.post("https://api.monday.com/v2", headers=monday_headers, json={"query": del_query})
                except Exception:
                    pass
        else:
            logger.info(f"Background Thread: No subitems found on duplicated GSA item {new_item_id} after waiting.")
            
        # Update column values and move group (Status: Contract Sent)
        update_query = """
        mutation ($itemId: ID!, $boardId: ID!, $group: String!, $colVals: JSON!) {
          move_item_to_group (item_id: $itemId, group_id: $group) { id }
          change_multiple_column_values (item_id: $itemId, board_id: $boardId, column_values: $colVals) { id }
        }
        """
        vals_dict = {
            "name": inn_code,
            "color_mm5zq7cg": {"label": "GSA ONLY"},
            "color_mm5w8ehb": {"label": "Gathering Info"},
            "status": {"label": "Sales"},
            "color_mm65xpj1": {"label": "Active"}
        }
        if gsa_quote_id:
            vals_dict["text_mm8v74nz"] = str(gsa_quote_id)
        if total_monthly is not None:
            vals_dict["numeric_mm8w91ec"] = total_monthly
            
        col_vals = json.dumps(vals_dict)
        
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
            logger.info(f"Successfully moved and updated GSA item {new_item_id}.")
            
            # Poll for newly automated subitems and mark 'Send GSA Contract' as Done
            target_subitem_name = "Send GSA Contract"
            subitem_to_update = None
            max_auto_retries = 15
            
            logger.info(f"Polling for automated subitems on GSA item {new_item_id}...")
            for attempt in range(max_auto_retries):
                sub_query = """
                query ($itemId: [ID!]) {
                  items (ids: $itemId) {
                    subitems { id name }
                  }
                }
                """
                sub_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                    "query": sub_query,
                    "variables": {"itemId": [new_item_id]}
                })
                try:
                    current_subitems = sub_resp.json().get('data', {}).get('items', [{}])[0].get('subitems', [])
                    # Check if the target subitem has been created
                    found_target = next((sub for sub in current_subitems if sub.get('name') == target_subitem_name), None)
                    if found_target:
                        subitem_to_update = found_target
                        break
                except Exception:
                    pass
                time.sleep(3.0)
                
            if subitem_to_update:
                logger.info(f"Found target subitem {target_subitem_name}. Marking as Done...")
                update_sub_query = """
                mutation ($itemId: ID!, $boardId: ID!, $colVals: JSON!) {
                  change_multiple_column_values (item_id: $itemId, board_id: $boardId, column_values: $colVals) { id }
                }
                """
                requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                    "query": update_sub_query,
                    "variables": {
                        "itemId": subitem_to_update['id'],
                        "boardId": "18424851639",
                        "colVals": json.dumps({"status": {"label": "Done"}})
                    }
                })
                
                # Now set main status to Contract Sent
                logger.info("Setting main item status to Contract Sent.")
                requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                    "query": update_sub_query,
                    "variables": {
                        "itemId": new_item_id,
                        "boardId": board_id,
                        "colVals": json.dumps({"color_mm6sy6mm": {"label": "Contract Sent"}})
                    }
                })
            else:
                logger.warning(f"Timed out waiting for automated subitem '{target_subitem_name}' to appear.")
                
    except Exception as e:
        logger.error(f"Failed in process_monday_gsa_duplicate_async: {e}")

def process_monday_cabling_duplicate_async(new_item_id, inn_code, monday_token, cabling_quote_num=None, grand_total=None):
    """Background task to clean up subitems and update status on a newly duplicated Cabling Monday item."""
    board_id = "18424851511"
    monday_headers = {
        "Authorization": monday_token, 
        "API-Version": "2024-01", 
        "Content-Type": "application/json"
    }

    try:
        # Clean up subitems from the duplicated item FIRST (before status changes trigger automations)
        subitems = []
        max_retries = 15
        last_count = -1
        stable_polls = 0
        
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
                current_subitems = sub_resp.json().get('data', {}).get('items', [{}])[0].get('subitems', [])
                current_count = len(current_subitems)
                
                if current_count > 0:
                    if current_count == last_count:
                        stable_polls += 1
                        if stable_polls >= 2:
                            subitems = current_subitems
                            break
                    else:
                        stable_polls = 0
                last_count = current_count
            except Exception:
                pass
            time.sleep(2.5)
            
        if subitems:
            logger.info(f"Background Thread: Found {len(subitems)} subitems on duplicated Cabling item {new_item_id}. Deleting them...")
            del_mutations = []
            for idx, sub in enumerate(subitems):
                del_mutations.append(f"del{idx}: delete_item (item_id: {sub['id']}) {{ id }}")
            
            chunk_size = 30
            for i in range(0, len(del_mutations), chunk_size):
                chunk = del_mutations[i:i + chunk_size]
                del_query = "mutation { " + " ".join(chunk) + " }"
                requests.post("https://api.monday.com/v2", headers=monday_headers, json={"query": del_query})
                
            # Verify deletion and delete any stragglers
            for _ in range(5):
                time.sleep(2.0)
                check_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
                    "query": sub_query,
                    "variables": {"itemId": [new_item_id]}
                })
                try:
                    remaining = check_resp.json().get('data', {}).get('items', [{}])[0].get('subitems', [])
                    if not remaining:
                        logger.info("Background Thread: Verified all subitems deleted.")
                        break
                    else:
                        logger.warning(f"Background Thread: {len(remaining)} subitems still remain. Retrying delete...")
                        del_mutations = []
                        for idx, sub in enumerate(remaining):
                            del_mutations.append(f"del{idx}: delete_item (item_id: {sub['id']}) {{ id }}")
                        
                        for i in range(0, len(del_mutations), chunk_size):
                            chunk = del_mutations[i:i + chunk_size]
                            del_query = "mutation { " + " ".join(chunk) + " }"
                            requests.post("https://api.monday.com/v2", headers=monday_headers, json={"query": del_query})
                except Exception:
                    pass
        else:
            logger.info(f"Background Thread: No subitems found on duplicated Cabling item {new_item_id} after waiting.")
            
        # Update column values (Status: Working on it)
        update_query = """
        mutation ($itemId: ID!, $boardId: ID!, $colVals: JSON!) {
          change_multiple_column_values (item_id: $itemId, board_id: $boardId, column_values: $colVals) { id }
        }
        """
        vals_dict = {
            "name": inn_code,
            "color_mm5zq7cg": {"label": "CABLING"},
            "color_mm5w8ehb": {"label": "Quote Sent to Customer"},
            "status": {"label": "Sales"},
            "color_mm65xpj1": {"label": "Active"},
            "color_mm6sy6mm": {"label": "Working on it"}
        }
        if cabling_quote_num and cabling_quote_num != "UNKNOWN":
            vals_dict["text_mm8v74nz"] = str(cabling_quote_num)
        if grand_total is not None:
            vals_dict["numeric_mm8vcbpj"] = grand_total
            
        col_vals = json.dumps(vals_dict)
        
        update_resp = requests.post("https://api.monday.com/v2", headers=monday_headers, json={
            "query": update_query,
            "variables": {
                "itemId": new_item_id,
                "boardId": board_id,
                "colVals": col_vals
            }
        })
        up_data = update_resp.json()
        if "errors" in up_data:
            logger.error(f"Monday Update Duplicated Cabling Item Error: {up_data['errors']}")
        else:
            logger.info(f"Successfully updated Cabling item {new_item_id}.")
                
    except Exception as e:
        logger.error(f"Failed in process_monday_cabling_duplicate_async: {e}")
