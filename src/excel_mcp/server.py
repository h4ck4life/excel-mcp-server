#!/usr/bin/env python
# server.py
import logging
import sys
import os
import asyncio
from typing import Any, List, Dict
import urllib.parse
from typing import Optional

from mcp.server import Server
from mcp.types import Resource, Tool, TextContent

# Import exceptions
from excel_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError
)

# Import from excel_mcp package with consistent _impl suffixes
from excel_mcp.validation import (
    validate_formula_in_cell_operation as validate_formula_impl,
    validate_range_in_sheet_operation as validate_range_impl
)
from excel_mcp.chart import create_chart_in_sheet as create_chart_impl
from excel_mcp.workbook import get_workbook_info
from excel_mcp.data import write_data
from excel_mcp.pivot import create_pivot_table as create_pivot_table_impl
from excel_mcp.sheet import (
    copy_sheet,
    delete_sheet,
    rename_sheet,
    merge_range,
    unmerge_range,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stderr),
        logging.FileHandler("excel-mcp.log")
    ],
    force=True
)

logger = logging.getLogger("excel-mcp")

# Get Excel files path from environment or use default
EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")

# Create the directory if it doesn't exist
os.makedirs(EXCEL_FILES_PATH, exist_ok=True)

# Initialize Server instead of FastMCP
app = Server("excel-mcp")

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
    """
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename
        
    # Use the configured Excel files path
    return os.path.join(EXCEL_FILES_PATH, filename)


@app.list_tools()
async def list_tools() -> list[Tool]:
    """List available Excel tools."""
    logger.info("Listing tools...")
    return [
        Tool(
            name="apply_formula",
            description="Apply Excel formula to cell",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "cell": {"type": "string", "description": "Cell reference (e.g., A1)"},
                    "formula": {"type": "string", "description": "Excel formula to apply"}
                },
                "required": ["filepath", "sheet_name", "cell", "formula"]
            }
        ),
        Tool(
            name="validate_formula_syntax",
            description="Validate Excel formula syntax without applying it",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "cell": {"type": "string", "description": "Cell reference (e.g., A1)"},
                    "formula": {"type": "string", "description": "Excel formula to validate"}
                },
                "required": ["filepath", "sheet_name", "cell", "formula"]
            }
        ),
        Tool(
            name="format_range",
            description="Apply formatting to a range of cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "start_cell": {"type": "string", "description": "Start cell of range"},
                    "end_cell": {"type": "string", "description": "End cell of range"},
                    "bold": {"type": "boolean", "description": "Apply bold formatting"},
                    "italic": {"type": "boolean", "description": "Apply italic formatting"},
                    "font_size": {"type": "integer", "description": "Font size"},
                    "font_color": {"type": "string", "description": "Font color"},
                    "bg_color": {"type": "string", "description": "Background color"}
                },
                "required": ["filepath", "sheet_name", "start_cell"]
            }
        ),
        Tool(
            name="read_data_from_excel",
            description="Read data from Excel worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "start_cell": {"type": "string", "description": "Start cell of range (default: A1)"},
                    "end_cell": {"type": "string", "description": "End cell of range"},
                    "preview_only": {"type": "boolean", "description": "Only show preview of data"}
                },
                "required": ["filepath", "sheet_name"]
            }
        ),
        Tool(
            name="write_data_to_excel",
            description="Write data to Excel worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "data": {"type": "array", "description": "List of rows to write"},
                    "start_cell": {"type": "string", "description": "Start cell (default: A1)"},
                    "write_headers": {"type": "boolean", "description": "Write headers"}
                },
                "required": ["filepath", "sheet_name", "data"]
            }
        ),
        Tool(
            name="create_workbook",
            description="Create new Excel workbook",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"}
                },
                "required": ["filepath"]
            }
        ),
        Tool(
            name="create_worksheet",
            description="Create new worksheet in workbook",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of new worksheet"}
                },
                "required": ["filepath", "sheet_name"]
            }
        ),
        Tool(
            name="create_chart",
            description="Create chart in worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "data_range": {"type": "string", "description": "Data range for chart"},
                    "chart_type": {"type": "string", "description": "Type of chart"},
                    "target_cell": {"type": "string", "description": "Target cell for chart"},
                    "title": {"type": "string", "description": "Chart title"},
                    "x_axis": {"type": "string", "description": "X-axis label"},
                    "y_axis": {"type": "string", "description": "Y-axis label"}
                },
                "required": ["filepath", "sheet_name", "data_range", "chart_type", "target_cell"]
            }
        ),
        Tool(
            name="create_pivot_table",
            description="Create pivot table in worksheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "data_range": {"type": "string", "description": "Data range for pivot table"},
                    "rows": {"type": "array", "description": "Row fields"},
                    "values": {"type": "array", "description": "Value fields"},
                    "columns": {"type": "array", "description": "Column fields"},
                    "agg_func": {"type": "string", "description": "Aggregation function"}
                },
                "required": ["filepath", "sheet_name", "data_range", "rows", "values"]
            }
        ),
        Tool(
            name="copy_worksheet",
            description="Copy worksheet within workbook",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "source_sheet": {"type": "string", "description": "Source worksheet name"},
                    "target_sheet": {"type": "string", "description": "Target worksheet name"}
                },
                "required": ["filepath", "source_sheet", "target_sheet"]
            }
        ),
        Tool(
            name="delete_worksheet",
            description="Delete worksheet from workbook",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet to delete"}
                },
                "required": ["filepath", "sheet_name"]
            }
        ),
        Tool(
            name="rename_worksheet",
            description="Rename worksheet in workbook",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "old_name": {"type": "string", "description": "Current worksheet name"},
                    "new_name": {"type": "string", "description": "New worksheet name"}
                },
                "required": ["filepath", "old_name", "new_name"]
            }
        ),
        Tool(
            name="get_workbook_metadata",
            description="Get metadata about workbook including sheets, ranges, etc.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "include_ranges": {"type": "boolean", "description": "Include named ranges"}
                },
                "required": ["filepath"]
            }
        ),
        Tool(
            name="merge_cells",
            description="Merge a range of cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "start_cell": {"type": "string", "description": "Start cell of range"},
                    "end_cell": {"type": "string", "description": "End cell of range"}
                },
                "required": ["filepath", "sheet_name", "start_cell", "end_cell"]
            }
        ),
        Tool(
            name="unmerge_cells",
            description="Unmerge a range of cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "start_cell": {"type": "string", "description": "Start cell of range"},
                    "end_cell": {"type": "string", "description": "End cell of range"}
                },
                "required": ["filepath", "sheet_name", "start_cell", "end_cell"]
            }
        ),
        Tool(
            name="copy_range",
            description="Copy a range of cells to another location",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Source worksheet name"},
                    "source_start": {"type": "string", "description": "Source start cell"},
                    "source_end": {"type": "string", "description": "Source end cell"},
                    "target_start": {"type": "string", "description": "Target start cell"},
                    "target_sheet": {"type": "string", "description": "Target worksheet name"}
                },
                "required": ["filepath", "sheet_name", "source_start", "source_end", "target_start"]
            }
        ),
        Tool(
            name="delete_range",
            description="Delete a range of cells and shift remaining cells",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "start_cell": {"type": "string", "description": "Start cell of range"},
                    "end_cell": {"type": "string", "description": "End cell of range"},
                    "shift_direction": {"type": "string", "description": "Direction to shift cells (up/left)"}
                },
                "required": ["filepath", "sheet_name", "start_cell", "end_cell"]
            }
        ),
        Tool(
            name="validate_excel_range",
            description="Validate if a range exists and is properly formatted",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to Excel file"},
                    "sheet_name": {"type": "string", "description": "Name of worksheet"},
                    "start_cell": {"type": "string", "description": "Start cell of range"},
                    "end_cell": {"type": "string", "description": "End cell of range"}
                },
                "required": ["filepath", "sheet_name", "start_cell"]
            }
        )
    ]

@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Call Excel tools."""
    logger.info(f"Calling tool: {name} with arguments: {arguments}")
    
    try:
        if name == "apply_formula":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            cell = arguments.get("cell", "")
            formula = arguments.get("formula", "")
            
            # Validate inputs
            if not all([filepath, sheet_name, cell, formula]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                # First validate the formula
                validation = validate_formula_impl(filepath, sheet_name, cell, formula)
                if isinstance(validation, dict) and "error" in validation:
                    return [TextContent(type="text", text=f"Error: {validation['error']}")]
                    
                # If valid, apply the formula
                from excel_mcp.calculations import apply_formula as apply_formula_impl
                result = apply_formula_impl(filepath, sheet_name, cell, formula)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, CalculationError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "validate_formula_syntax":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            cell = arguments.get("cell", "")
            formula = arguments.get("formula", "")
            
            # Validate inputs
            if not all([filepath, sheet_name, cell, formula]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = validate_formula_impl(filepath, sheet_name, cell, formula)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, CalculationError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "format_range":
            # Handle formatting tool
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            start_cell = arguments.get("start_cell", "")
            
            if not all([filepath, sheet_name, start_cell]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                from excel_mcp.formatting import format_range as format_range_func
                
                result = format_range_func(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    start_cell=start_cell,
                    end_cell=arguments.get("end_cell"),
                    bold=arguments.get("bold", False),
                    italic=arguments.get("italic", False),
                    underline=arguments.get("underline", False),
                    font_size=arguments.get("font_size"),
                    font_color=arguments.get("font_color"),
                    bg_color=arguments.get("bg_color"),
                    border_style=arguments.get("border_style"),
                    border_color=arguments.get("border_color"),
                    number_format=arguments.get("number_format"),
                    alignment=arguments.get("alignment"),
                    wrap_text=arguments.get("wrap_text", False),
                    merge_cells=arguments.get("merge_cells", False),
                    protection=arguments.get("protection"),
                    conditional_format=arguments.get("conditional_format")
                )
                return [TextContent(type="text", text="Range formatted successfully")]
            except (ValidationError, FormattingError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "read_data_from_excel":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            start_cell = arguments.get("start_cell", "A1")
            end_cell = arguments.get("end_cell")
            preview_only = arguments.get("preview_only", False)
            
            if not all([filepath, sheet_name]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                from excel_mcp.data import read_excel_range
                result = read_excel_range(filepath, sheet_name, start_cell, end_cell, preview_only)
                if not result:
                    return [TextContent(type="text", text="No data found in specified range")]
                # Convert the list of dicts to a formatted string
                data_str = "\n".join([str(row) for row in result])
                return [TextContent(type="text", text=data_str)]
            except Exception as e:
                return [TextContent(type="text", text=f"Error reading data: {str(e)}")]
                
        elif name == "write_data_to_excel":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            data = arguments.get("data", [])
            start_cell = arguments.get("start_cell", "A1")
            write_headers = arguments.get("write_headers", True)
            
            if not all([filepath, sheet_name, data]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = write_data(filepath, sheet_name, data, start_cell, write_headers)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, DataError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "create_workbook":
            filepath = get_excel_path(arguments.get("filepath", ""))
            
            if not filepath:
                return [TextContent(type="text", text="Error: Missing filepath parameter")]
                
            try:
                from excel_mcp.workbook import create_workbook as create_workbook_impl
                result = create_workbook_impl(filepath)
                return [TextContent(type="text", text=f"Created workbook at {filepath}")]
            except WorkbookError as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "create_worksheet":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            
            if not all([filepath, sheet_name]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                from excel_mcp.workbook import create_sheet as create_worksheet_impl
                result = create_worksheet_impl(filepath, sheet_name)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, WorkbookError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "create_chart":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            data_range = arguments.get("data_range", "")
            chart_type = arguments.get("chart_type", "")
            target_cell = arguments.get("target_cell", "")
            title = arguments.get("title", "")
            x_axis = arguments.get("x_axis", "")
            y_axis = arguments.get("y_axis", "")
            
            if not all([filepath, sheet_name, data_range, chart_type, target_cell]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = create_chart_impl(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    data_range=data_range,
                    chart_type=chart_type,
                    target_cell=target_cell,
                    title=title,
                    x_axis=x_axis,
                    y_axis=y_axis
                )
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, ChartError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "create_pivot_table":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            data_range = arguments.get("data_range", "")
            rows = arguments.get("rows", [])
            values = arguments.get("values", [])
            columns = arguments.get("columns", [])
            agg_func = arguments.get("agg_func", "mean")
            
            if not all([filepath, sheet_name, data_range, rows, values]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = create_pivot_table_impl(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    data_range=data_range,
                    rows=rows,
                    values=values,
                    columns=columns or [],
                    agg_func=agg_func
                )
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, PivotError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "copy_worksheet":
            filepath = get_excel_path(arguments.get("filepath", ""))
            source_sheet = arguments.get("source_sheet", "")
            target_sheet = arguments.get("target_sheet", "")
            
            if not all([filepath, source_sheet, target_sheet]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = copy_sheet(filepath, source_sheet, target_sheet)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "delete_worksheet":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            
            if not all([filepath, sheet_name]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = delete_sheet(filepath, sheet_name)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "rename_worksheet":
            filepath = get_excel_path(arguments.get("filepath", ""))
            old_name = arguments.get("old_name", "")
            new_name = arguments.get("new_name", "")
            
            if not all([filepath, old_name, new_name]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = rename_sheet(filepath, old_name, new_name)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "get_workbook_metadata":
            filepath = get_excel_path(arguments.get("filepath", ""))
            include_ranges = arguments.get("include_ranges", False)
            
            if not filepath:
                return [TextContent(type="text", text="Error: Missing filepath parameter")]
                
            try:
                result = get_workbook_info(filepath, include_ranges=include_ranges)
                return [TextContent(type="text", text=str(result))]
            except WorkbookError as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "merge_cells":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            start_cell = arguments.get("start_cell", "")
            end_cell = arguments.get("end_cell", "")
            
            if not all([filepath, sheet_name, start_cell, end_cell]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = merge_range(filepath, sheet_name, start_cell, end_cell)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "unmerge_cells":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            start_cell = arguments.get("start_cell", "")
            end_cell = arguments.get("end_cell", "")
            
            if not all([filepath, sheet_name, start_cell, end_cell]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                result = unmerge_range(filepath, sheet_name, start_cell, end_cell)
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "copy_range":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            source_start = arguments.get("source_start", "")
            source_end = arguments.get("source_end", "")
            target_start = arguments.get("target_start", "")
            target_sheet = arguments.get("target_sheet")
            
            if not all([filepath, sheet_name, source_start, source_end, target_start]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                from excel_mcp.sheet import copy_range_operation
                result = copy_range_operation(
                    filepath,
                    sheet_name,
                    source_start,
                    source_end,
                    target_start,
                    target_sheet
                )
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "delete_range":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            start_cell = arguments.get("start_cell", "")
            end_cell = arguments.get("end_cell", "")
            shift_direction = arguments.get("shift_direction", "up")
            
            if not all([filepath, sheet_name, start_cell, end_cell]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                from excel_mcp.sheet import delete_range_operation
                result = delete_range_operation(
                    filepath,
                    sheet_name,
                    start_cell,
                    end_cell,
                    shift_direction
                )
                return [TextContent(type="text", text=result["message"])]
            except (ValidationError, SheetError) as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
                
        elif name == "validate_excel_range":
            filepath = get_excel_path(arguments.get("filepath", ""))
            sheet_name = arguments.get("sheet_name", "")
            start_cell = arguments.get("start_cell", "")
            end_cell = arguments.get("end_cell")
            
            if not all([filepath, sheet_name, start_cell]):
                return [TextContent(type="text", text="Error: Missing required parameters")]
                
            try:
                range_str = start_cell if not end_cell else f"{start_cell}:{end_cell}"
                result = validate_range_impl(filepath, sheet_name, range_str)
                return [TextContent(type="text", text=result["message"])]
            except ValidationError as e:
                return [TextContent(type="text", text=f"Error: {str(e)}")]
        
        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]
            
    except Exception as e:
        logger.error(f"Error calling tool {name}: {e}", exc_info=True)
        return [TextContent(type="text", text=f"Error executing tool: {str(e)}")]

@app.list_resources()
async def list_resources() -> list[Resource]:
    """List Excel files as resources."""
    try:
        files = []
        for file in os.listdir(EXCEL_FILES_PATH):
            if file.endswith(('.xlsx', '.xls', '.xlsm')):
                # URL encode the filename to make it URL-safe
                encoded_file = urllib.parse.quote(file)
                # Make sure to use proper format for resource
                files.append(
                    Resource(
                        uri=f"excel://{encoded_file}",
                        name=file, 
                        mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        description=f"Excel workbook: {file}"
                    )
                )
        return files
    except Exception as e:
        logger.error(f"Failed to list resources: {str(e)}")
        # Return empty list instead of raising exception
        return []

@app.read_resource()
async def read_resource(uri: str) -> str:
    """Read Excel workbook contents."""
    try:
        if not uri.startswith("excel://"):
            raise ValueError(f"Invalid URI scheme: {uri}")
            
        # URL decode the filename part
        encoded_file = uri[8:]  # Remove "excel://" prefix
        file = urllib.parse.unquote(encoded_file)
        
        file_path = get_excel_path(file)
        
        # Check if file exists first
        if not os.path.exists(file_path):
            return f"Error: File not found: {file}"
        
        # Get workbook info with timeout handling
        from concurrent.futures import ThreadPoolExecutor
        import asyncio
        
        def get_info():
            try:
                return get_workbook_info(file_path, include_ranges=True)
            except Exception as e:
                return f"Error reading workbook: {str(e)}"
        
        # Run potentially slow operation with timeout
        with ThreadPoolExecutor() as executor:
            try:
                result = await asyncio.wait_for(
                    asyncio.get_event_loop().run_in_executor(executor, get_info),
                    timeout=15.0  # 15 second timeout
                )
                return str(result)
            except asyncio.TimeoutError:
                logger.error(f"Timeout while reading resource {uri}")
                return "Error: Operation timed out while reading Excel file"
    except Exception as e:
        logger.error(f"Error reading resource {uri}: {e}")
        return f"Error reading resource: {str(e)}"

async def main():
    """Main entry point to run the MCP server."""
    from mcp.server.stdio import stdio_server
    
    logger.info("Starting Excel MCP server...")
    try:
        logger.info(f"Excel files directory: {EXCEL_FILES_PATH}")
        
        # Use async with pattern for cleaner resource management
        async with stdio_server() as (read_stream, write_stream):
            try:
                # Add a short delay for stability
                await asyncio.sleep(1)
                
                # Run the server with standard initialization options
                await app.run(
                    read_stream,
                    write_stream,
                    app.create_initialization_options()
                )
            except Exception as e:
                logger.error(f"Server error: {str(e)}", exc_info=True)
                # Allow a moment for cleanup before potential restart
                await asyncio.sleep(1)
    except Exception as e:
        logger.error(f"Startup error: {str(e)}", exc_info=True)
    finally:
        logger.info("Excel MCP server stopped")

# This makes it work with 'uv run excel-mcp-server'
def run_server():
    """Function to be called by uv run."""
    asyncio.run(main())

# Also keep traditional entry point for direct script execution
if __name__ == "__main__":
    asyncio.run(main())