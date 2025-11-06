import json

def generate_embedded_data():
    """Generate embedded JavaScript files from JSON data to avoid CORS issues."""
    
    # Read MDB data
    try:
        with open('mdb_data.json', 'r') as f:
            mdb_data = json.load(f)
    except FileNotFoundError:
        print("Error: mdb_data.json not found. Run extract_mdb_data.py first.")
        return
    
    # Read all boards data
    try:
        with open('all_boards_data.json', 'r') as f:
            all_boards_data = json.load(f)
    except FileNotFoundError:
        print("Warning: all_boards_data.json not found. Using MDB data only.")
        all_boards_data = {
            'all_boards': mdb_data.get('mdb_boards', []),
            'total_estimate': mdb_data.get('total_estimate', 0),
            'count': mdb_data.get('count', 0)
        }
    
    # Generate embedded data file
    # Use compact JSON format to reduce file size
    embedded_content = f"""// Embedded dashboard data - Generated automatically
// This file contains the data embedded to avoid CORS issues when opening HTML directly
// Generated from e2.xlsx TOTALLIST sheet
// To regenerate: python3 generate_embedded_data.py

window.dashboardData = {{
    mdbBoards: {json.dumps(mdb_data.get('mdb_boards', []), separators=(',', ':'))},
    totalEstimate: {mdb_data.get('total_estimate', 0)},
    totalLoad: {mdb_data.get('total_load', 0)},
    totalItems: {mdb_data.get('total_items', 0)},
    mdbCount: {mdb_data.get('count', 0)},
    allBoards: {json.dumps(all_boards_data.get('all_boards', []), separators=(',', ':'))},
    allBoardsTotal: {all_boards_data.get('total_estimate', 0)},
    allBoardsTotalLoad: {all_boards_data.get('total_load', 0)},
    allBoardsTotalItems: {all_boards_data.get('total_items', 0)},
    allBoardsCount: {all_boards_data.get('count', 0)}
}};
"""
    
    # Write to file
    with open('embed_data.js', 'w') as f:
        f.write(embedded_content)
    
    print("✓ Generated embed_data.js successfully")
    print(f"  - Main MDBs: {mdb_data.get('count', 0)}")
    print(f"  - Total MDB Estimate: {mdb_data.get('total_estimate', 0):,.2f}")
    print(f"  - All Boards: {all_boards_data.get('count', 0)}")
    print(f"  - Total All Boards Estimate: {all_boards_data.get('total_estimate', 0):,.2f}")

if __name__ == '__main__':
    generate_embedded_data()

