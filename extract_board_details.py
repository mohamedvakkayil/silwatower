import openpyxl
import json
import sys

def extract_board_details(board_name):
    """Extract detailed data from a specific board sheet."""
    try:
        wb = openpyxl.load_workbook('e2.xlsx', data_only=True)
        
        if board_name not in wb.sheetnames:
            return {'error': f'Sheet "{board_name}" not found'}
        
        ws = wb[board_name]
        
        # Extract all data from the sheet
        details = {
            'board_name': board_name,
            'items': [],
            'summary': {}
        }
        
        # Find header row (usually row 1 or 2)
        header_row = None
        for row_idx in range(1, min(10, ws.max_row + 1)):
            row = ws[row_idx]
            # Check if this looks like a header row (contains common headers)
            if any(cell and isinstance(cell.value, str) and 
                   any(header in str(cell.value).upper() for header in ['ITEM', 'DESCRIPTION', 'QTY', 'UNIT', 'RATE', 'AMOUNT'])
                   for cell in row[:10]):
                header_row = row_idx
                break
        
        if header_row is None:
            header_row = 1
        
        # Extract headers
        headers = []
        for col in range(1, min(ws.max_column + 1, 20)):
            cell_value = ws.cell(header_row, col).value
            if cell_value:
                headers.append(str(cell_value).strip())
            else:
                headers.append(f'Column{col}')
        
        # Extract data rows
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_data = {}
            has_data = False
            
            for col_idx, header in enumerate(headers):
                if col_idx < ws.max_column:
                    cell_value = ws.cell(row_idx, col_idx + 1).value
                    if cell_value is not None:
                        row_data[header] = cell_value
                        has_data = True
            
            if has_data:
                details['items'].append(row_data)
        
        # Find summary values (NET TOTAL, NO OF UNITS, etc.)
        for row_idx in range(max(1, ws.max_row - 30), ws.max_row + 1):
            row = ws[row_idx]
            for col_idx in range(min(6, ws.max_column)):
                cell_value = row[col_idx].value
                if cell_value and isinstance(cell_value, str):
                    cell_upper = str(cell_value).upper()
                    # Look for summary rows
                    if 'NET TOTAL' in cell_upper or 'TOTAL' in cell_upper:
                        # Try to get the value from amount column (usually column F, index 5)
                        if ws.max_column > 5:
                            amount_value = ws.cell(row_idx, 6).value
                            if amount_value is not None:
                                try:
                                    details['summary']['net_total'] = float(amount_value)
                                except (ValueError, TypeError):
                                    pass
                    elif 'NO OF UNITS' in cell_upper or 'NO OF ITEMS' in cell_upper:
                        if ws.max_column > 5:
                            units_value = ws.cell(row_idx, 6).value
                            if units_value is not None:
                                try:
                                    details['summary']['no_of_units'] = float(units_value)
                                except (ValueError, TypeError):
                                    pass
        
        return details
        
    except Exception as e:
        return {'error': str(e)}

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({'error': 'Board name required'}))
        sys.exit(1)
    
    board_name = sys.argv[1]
    result = extract_board_details(board_name)
    print(json.dumps(result, indent=2, default=str))

