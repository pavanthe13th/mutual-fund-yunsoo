"""
Table Structure Parser Module

This module provides functionality to identify and process table structures
from Excel sheets based on metadata configurations.
"""

import json
import logging
import math
import os
import re
from typing import Dict, List, Optional, Tuple, Any

logger = logging.getLogger(__name__)


def normalize_header_text(text: str) -> str:
    """
    Normalize header text for comparison by:
    1. Converting to lowercase
    2. Replacing newlines and multiple whitespace with single space
    3. Stripping footnote markers (~, ^, *)

    Args:
        text: Raw header text

    Returns:
        Normalized text for comparison
    """
    # Convert to lowercase and strip
    text = str(text).strip().lower()
    # Replace newlines and multiple whitespace with single space
    text = re.sub(r'\s+', ' ', text)
    # Remove common footnote markers at end of string
    text = re.sub(r'[~^*]+$', '', text)
    return text.strip()


def normalize_instrument_type(text: Optional[str]) -> Optional[str]:
    """
    Normalize instrument_type strings by stripping whitespace and collapsing multiple spaces.
    
    This ensures consistent formatting regardless of source (metadata JSON, Excel files, etc.)
    and prevents issues with extra spaces in instrument_type values.
    
    Args:
        text: The instrument_type string to normalize (can be None)
    
    Returns:
        Normalized string with single spaces, or None if input was None/empty
    """
    if not text:
        return None
    
    # Convert to string, strip leading/trailing whitespace, and collapse multiple spaces
    normalized = re.sub(r'\s+', ' ', str(text).strip())
    return normalized if normalized else None


class TableStructure:
    """Represents a table structure configuration."""

    def __init__(self, config: Dict):
        self.name = config['name']
        self.header_columns = config['header_columns']
        # Controls whether alternate_input_names are considered in header matching.
        self.use_alternate_input_names = False
        self.vertical_hierarchy = config.get('vertical_hierarchy', [])
        # Support both string and list for net_receivables_text
        net_recv = config.get('net_receivables_text', None)
        if net_recv is None:
            self.net_receivables_texts = []
        elif isinstance(net_recv, list):
            self.net_receivables_texts = net_recv
        else:
            self.net_receivables_texts = [net_recv]
        self.table_end_keywords = config.get('table_end_keywords', [])
        # If true, check ALL cells in row for table_end_keywords (not just first cell)
        # Useful when footer/disclaimer text appears in middle columns
        self.table_end_check_all_cells = config.get('table_end_check_all_cells', False)
        # Default instrument_type for all records in this table (overrides hierarchy)
        # Normalize to remove extra spaces
        self.default_instrument_type = normalize_instrument_type(
            config.get('default_instrument_type', None)
        )
        # Override instrument_type specifically for net_receivables rows (empty string "" means blank)
        # Don't normalize - preserve empty strings as explicit blank values
        self.net_receivables_instrument_type = config.get('net_receivables_instrument_type', None)
        # Override category specifically for net_receivables rows (empty string "" means blank)
        self.net_receivables_category = config.get('net_receivables_category', None)
        # Column index to check for hierarchy markers (default: check all columns)
        self.hierarchy_marker_column = config.get('hierarchy_marker_column', None)
        # List of instrument names for which instrument_type and category should be blank
        self.blank_hierarchy_instruments = [x.lower() for x in config.get('blank_hierarchy_instruments', [])]
        # Columns to exclude from output (but still used for header matching and validation)
        self.exclude_output_columns = config.get('exclude_output_columns', [])
        # Track which columns actually matched during header detection
        # This prevents extracting data from wrong columns when headers don't match
        self.matched_columns_set: set = set()
        # Patterns to exclude from instrument names (footnotes, disclaimers, etc.)
        self.exclude_instrument_patterns = config.get('exclude_instrument_patterns', [])
        # Close table when hitting exclude patterns after a TOTAL row (for tables without GRAND TOTAL)
        self.close_table_after_total_on_exclude_match = config.get('close_table_after_total_on_exclude_match', False)

    def _get_expected_header_names(self, col_def: Dict[str, Any]) -> List[str]:
        """Get accepted header names for a column, optionally including alternates."""
        expected_names = [col_def.get('input_name', '')]
        if self.use_alternate_input_names:
            expected_names.extend(col_def.get('alternate_input_names', []))
        # Preserve order while deduplicating.
        return list(dict.fromkeys(expected_names))

    def matches_header(self, row: List[Any]) -> bool:
        """
        Check if a row matches this table structure's header.

        Per specification: "the pattern match for the header row should check to see if
        the string in the cell contains the keyword for the header column. Not every column
        needs to be present for the header to be a match, as long as a majority of the row
        is a match with a header structure, we can start processing the table."

        Args:
            row: List of cell values from a row

        Returns:
            True if a majority of header columns match (cell contains the keyword)
        """
        # Group columns by (index, output_name) - multiple input_names are alternatives
        logical_columns = {}
        for col_def in self.header_columns:
            key = (col_def['index'], col_def['output_name'])
            if key not in logical_columns:
                logical_columns[key] = []
            logical_columns[key].extend(self._get_expected_header_names(col_def))

        matches = 0
        total_columns = len(logical_columns)

        # Check each logical column - match if ANY input_name matches
        for (col_index, _), input_names in logical_columns.items():
            # Check if index is within row bounds
            if col_index >= len(row):
                continue

            actual_value = row[col_index]

            # Handle None/NaN values - allow match if empty string is a valid input_name
            if actual_value is None or (isinstance(actual_value, float) and str(actual_value) == 'nan'):
                if '' in input_names:
                    matches += 1
                continue

            # Normalize actual value once
            actual_str = normalize_header_text(actual_value)

            # Check if ANY of the alternative input_names match
            for expected_keyword in input_names:
                expected_str = normalize_header_text(expected_keyword)
                if expected_str in actual_str:
                    matches += 1
                    break  # Found a match, no need to check other alternatives

        # Return True if majority of columns match
        return matches > (total_columns / 2)

    def get_header_match_details(self, row: List[Any]) -> Dict[str, Any]:
        """
        Get detailed information about header matching for validation purposes.

        Args:
            row: List of cell values from a row

        Returns:
            Dictionary with match details:
            - is_match: True if majority matched
            - matched_count: Number of columns matched
            - total_count: Total columns expected
            - match_percentage: Percentage of columns matched
            - matched_columns: List of column names that matched
            - missing_columns: List of column names that didn't match
        """
        # Group columns by (index, output_name) - multiple input_names are alternatives
        logical_columns = {}
        for col_def in self.header_columns:
            key = (col_def['index'], col_def['output_name'])
            if key not in logical_columns:
                logical_columns[key] = []
            logical_columns[key].extend(self._get_expected_header_names(col_def))

        matched_columns = []
        missing_columns = []
        total_columns = len(logical_columns)

        for (col_index, output_name), input_names in logical_columns.items():
            # Use output_name for reporting (cleaner than listing all alternatives)
            col_label = output_name

            # Check if index is within row bounds
            if col_index >= len(row):
                missing_columns.append(col_label)
                continue

            actual_value = row[col_index]

            # Handle None/NaN values - allow match if empty string is a valid input_name
            if actual_value is None or (isinstance(actual_value, float) and str(actual_value) == 'nan'):
                if '' in input_names:
                    matched_columns.append(col_label)
                else:
                    missing_columns.append(col_label)
                continue

            # Normalize actual value once
            actual_str = normalize_header_text(actual_value)

            # Check if ANY of the alternative input_names match
            found_match = False
            for expected_keyword in input_names:
                expected_str = normalize_header_text(expected_keyword)
                if expected_str in actual_str:
                    matched_columns.append(col_label)
                    found_match = True
                    break

            if not found_match:
                missing_columns.append(col_label)

        matched_count = len(matched_columns)
        match_percentage = (matched_count / total_columns * 100) if total_columns > 0 else 0
        is_match = matched_count > (total_columns / 2)

        # Track unrecognized columns (in Excel but not in metadata)
        # These are non-empty columns that weren't matched to any metadata column
        unrecognized_columns = []
        matched_indices = {col_def['index'] for col_def in self.header_columns}
        for idx, cell in enumerate(row):
            if idx not in matched_indices:
                if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan'):
                    cell_str = str(cell).strip()
                    # Skip short strings, numbers, and common non-header patterns
                    if cell_str and len(cell_str) > 2:
                        # Skip if it's a number (row index, serial number)
                        try:
                            float(cell_str)
                            continue
                        except ValueError:
                            pass
                        unrecognized_columns.append(cell_str)

        return {
            'is_match': is_match,
            'matched_count': matched_count,
            'total_count': total_columns,
            'match_percentage': match_percentage,
            'matched_columns': matched_columns,
            'missing_columns': missing_columns,
            'unrecognized_columns': unrecognized_columns
        }

    def extract_record(self, row: List[Any]) -> Dict[str, Any]:
        """
        Extract a data record from a row based on column configuration.

        Only extracts data from columns that were successfully matched during
        header detection. This prevents incorrectly parsing data from wrong
        columns when the header doesn't match (e.g., different column at that index).

        Args:
            row: List of cell values from a row

        Returns:
            Dictionary with extracted field values
        """
        record = {}

        for col_def in self.header_columns:
            col_name = col_def['output_name']
            col_index = col_def['index']

            # Skip columns that didn't match during header detection
            # This prevents extracting data from wrong columns when header is different
            if self.matched_columns_set and col_name not in self.matched_columns_set:
                record[col_name] = None
                continue

            # Check if merge_indices is specified (merge multiple columns)
            if 'merge_indices' in col_def:
                values = []
                for idx in col_def['merge_indices']:
                    if idx < len(row) and row[idx] is not None:
                        val = str(row[idx]).strip()
                        if val and val.lower() != 'nan':
                            # Convert decimal coupon rate to percentage format
                            # e.g., 0.0925 -> "9.25%"
                            val = self._format_coupon_if_applicable(val)
                            values.append(val)
                value_str = ' '.join(values) if values else None
                record[col_name] = value_str
            else:
                # Original single-column logic
                if col_index < len(row):
                    value = row[col_index]
                    # Clean the value
                    if value is not None:
                        # Preserve numeric types instead of converting everything to string
                        # This prevents precision loss for very small numbers and avoids
                        # Excel storing them as text, which can cause parsing issues
                        if isinstance(value, (int, float)):
                            # Check for NaN (float('nan') comparison)
                            if isinstance(value, float) and math.isnan(value):
                                record[col_name] = None
                            else:
                                record[col_name] = value  # Keep as number
                        else:
                            # For strings and other types, convert to string
                            value_str = str(value).strip()
                            if value_str and value_str.lower() != 'nan':
                                record[col_name] = value_str
                            else:
                                record[col_name] = None
                    else:
                        record[col_name] = None
                else:
                    record[col_name] = None

        return record

    def _format_coupon_if_applicable(self, val: str) -> str:
        """
        Convert decimal coupon rates to percentage format.

        Args:
            val: String value that might be a decimal coupon rate

        Returns:
            Formatted string - percentage if it's a coupon rate, otherwise unchanged
        """
        try:
            num = float(val)
            # Check if it looks like a decimal coupon rate (between 0 and 1, exclusive)
            # Typical coupon rates are 0.05 to 0.15 (5% to 15%)
            if 0 < num < 1:
                # Convert to percentage and format nicely
                percentage = num * 100
                # Format to 2 decimal places, remove trailing zeros
                formatted = f"{percentage:.2f}".rstrip('0').rstrip('.')
                return f"{formatted}%"
            else:
                return val
        except (ValueError, TypeError):
            # Not a number, return as-is (e.g., "ZCB", "FRB")
            return val


class VerticalHierarchyTracker:
    """Tracks the current position in a vertical hierarchy.

    Supports 3-level hierarchy:
    - Level 1: Type of the Instrument
    - Level 2: Category
    - Level 3: SubCategory
    """

    def __init__(self, hierarchy_config: List[Dict], hierarchy_marker_column: Optional[int] = None, default_instrument_type: Optional[str] = None):
        self.hierarchy_config = hierarchy_config
        self.hierarchy_marker_column = hierarchy_marker_column  # Column index to check for hierarchy markers
        self.current_path: List[Dict] = []  # Stack of active hierarchy levels
        self.category_labels: List[str] = []  # Labels for current hierarchy path
        self.default_instrument_type = default_instrument_type

        # If default_instrument_type is set, pre-enter the instrument_type level
        # This allows category_type markers to be recognized without needing the parent instrument_type marker
        if default_instrument_type and hierarchy_config:
            # Create a synthetic entry for the instrument_type level
            # Find a matching config or use the first one's children as the active context
            for config in hierarchy_config:
                if config.get('children'):
                    # Pre-enter this instrument_type level so children (category_type) can be matched
                    self.current_path = [config]
                    self.category_labels = [default_instrument_type]
                    break

        # Validate that no instrument types contain overlapping substrings
        self._validate_no_substring_conflicts()

    def _validate_no_substring_conflicts(self):
        """
        Validate metadata for any issues.

        Since we use exact matching (after normalizing prefixes like "(a)"),
        substring conflicts are no longer an issue. This method is kept for
        future validation needs.
        """
        # No validation needed for exact matching
        pass

    def _normalize_cell_text(self, text: str) -> str:
        """
        Normalize cell text by stripping trailing spaces and prefixes like (a), a), i), etc.

        Args:
            text: The raw cell text

        Returns:
            Normalized text, or empty string for single-character non-alphanumeric values
        """
        import re
        # Strip whitespace
        text = text.strip()
        # Skip single-character non-alphanumeric values (like "|" pipe characters used as visual markers)
        # These are never valid hierarchy markers or meaningful text
        if len(text) == 1 and not text.isalnum():
            return ''
        # Remove prefix patterns:
        # - "(a) " or "(a)" - letter in parentheses
        # - "(db) " or "(xyz)" - multi-letter marker in parentheses (e.g., "(db) Government Securities")
        # - "a) " or "a)" - letter with closing paren only
        # - "i) " or "i)" - Roman numeral with closing paren
        text = re.sub(r'^\([a-z]+\)\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'^[a-z]\)\s*', '', text, flags=re.IGNORECASE)
        text = re.sub(r'^[ivxlcdm]+\)\s*', '', text, flags=re.IGNORECASE)
        return text

    def check_and_update(self, row: List[Any], data_column_indices: Optional[List[int]] = None, instrument_name_col_idx: Optional[int] = None) -> Tuple[bool, Optional[str]]:
        """
        Check if the row represents a hierarchy marker and update state.

        Uses exact matching after normalizing cell text (stripping prefixes like "(a)" and whitespace).

        Supports two modes:
        1. With end_keywords: Components close only when their end_keyword is encountered
        2. Without end_keywords: Components close on blank rows or sibling/parent start

        Args:
            row: List of cell values from a row
            data_column_indices: Optional list of column indices that contain data (ISIN, quantity, market_value, etc.)
                                If provided and row has data in these columns, it's treated as a data row, not a marker
            instrument_name_col_idx: Optional column index for instrument_name. If first cell doesn't match a marker,
                                    also check this column for hierarchy markers.

        Returns:
            Tuple of (is_hierarchy_marker, marker_type)
            - is_hierarchy_marker: True if this row is a hierarchy marker
            - marker_type: Type of marker ('start', 'end', or None)
        """
        # Determine which cell(s) to check for hierarchy markers
        normalized_cell = None
        raw_cell = None
        first_cell_idx = None

        if self.hierarchy_marker_column is not None:
            # Check specific column for hierarchy markers
            if self.hierarchy_marker_column < len(row):
                cell = row[self.hierarchy_marker_column]
                if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan') and str(cell).strip():
                    cell_str = str(cell)
                    raw_cell = cell_str.strip()
                    normalized = self._normalize_cell_text(cell_str)
                    if normalized:
                        normalized_cell = normalized
                        first_cell_idx = self.hierarchy_marker_column
        else:
            # Default: get the first non-empty cell, but skip cells that are only prefixes
            # (e.g., "a)", "i)") that normalize to empty strings
            for idx, cell in enumerate(row):
                if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan') and str(cell).strip():
                    cell_str = str(cell)
                    raw_cell = cell_str.strip()
                    normalized = self._normalize_cell_text(cell_str)
                    # If normalization resulted in a non-empty string, use it
                    if normalized:
                        normalized_cell = normalized
                        first_cell_idx = idx
                        break
                    # Otherwise, skip this cell and continue to the next

        if not normalized_cell:
            return False, None

        # Try matching with first cell
        result = self._match_hierarchy_marker(normalized_cell, raw_cell, data_column_indices, first_cell_idx, row)
        if result[0]:
            return result

        # If first cell doesn't match and instrument_name column is specified, check that column
        if instrument_name_col_idx is not None and instrument_name_col_idx != first_cell_idx and instrument_name_col_idx < len(row):
            cell = row[instrument_name_col_idx]
            if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan') and str(cell).strip():
                instrument_raw = str(cell).strip()
                instrument_normalized = self._normalize_cell_text(str(cell))
                if instrument_normalized:
                    return self._match_hierarchy_marker(instrument_normalized, instrument_raw, data_column_indices, instrument_name_col_idx, row)

        return False, None

    def _match_hierarchy_marker(self, normalized_cell: str, raw_cell: str, data_column_indices: Optional[List[int]], cell_idx: int, row: List[Any]) -> Tuple[bool, Optional[str]]:
        """Match a cell value against hierarchy markers.

        Maintains child-first matching order (Level 2/3 before Level 1) to prevent
        parent matches from overriding when inside a hierarchy.
        """
        # If data columns specified, check if row has data in those columns
        # Data rows should only check for end_keywords, not start markers
        if data_column_indices:
            for col_idx in data_column_indices:
                if col_idx < len(row) and col_idx != cell_idx:
                    cell = row[col_idx]
                    if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan'):
                        cell_str = str(cell).strip()
                        if cell_str and cell_str.lower() not in ['nil', 'n/a', 'na', '-', '']:
                            # Row has data - check for end_keywords but don't treat as start marker
                            # Exception: hierarchy markers with extract_data_from_marker can have data
                            for level_idx in range(len(self.current_path) - 1, -1, -1):
                                current_level = self.current_path[level_idx]
                                end_keywords = current_level.get('end_keywords', [])
                                if end_keywords:
                                    for keyword in end_keywords:
                                        if keyword.lower() in raw_cell.lower():
                                            while len(self.current_path) > level_idx:
                                                self.current_path.pop()
                                                if self.category_labels:
                                                    self.category_labels.pop()
                                            logger.debug(f"End keyword '{keyword}' matched in data row, closed level {level_idx + 1}")
                                            return True, 'end'

                            # Before skipping, check if this matches a hierarchy with extract_data_from_marker
                            for hierarchy_level in self.hierarchy_config:
                                instrument_type = hierarchy_level.get('instrument_type', '')
                                normalized_instrument = self._normalize_cell_text(instrument_type) if instrument_type else ''
                                if (normalized_instrument and
                                    normalized_cell.lower() == normalized_instrument.lower() and
                                    hierarchy_level.get('extract_data_from_marker', False)):
                                    # Allow this to proceed as a start marker
                                    break
                            else:
                                return False, None

        # STEP 1: Check for end_keywords FIRST (if present)
        # Check from deepest level backwards to close the most specific component first
        for level_idx in range(len(self.current_path) - 1, -1, -1):
            current_level = self.current_path[level_idx]
            end_keywords = current_level.get('end_keywords', [])

            if end_keywords:
                # Check if raw_cell contains any of the end keywords
                for keyword in end_keywords:
                    if keyword.lower() in raw_cell.lower():
                        # Found an end marker - pop this level and all deeper levels
                        while len(self.current_path) > level_idx:
                            self.current_path.pop()
                            if self.category_labels:
                                self.category_labels.pop()
                        logger.debug(f"End keyword '{keyword}' matched, closed level {level_idx + 1}")
                        return True, 'end'

        # When already inside a hierarchy, check for child matches FIRST
        # This ensures "B) MONEY MARKET INSTRUMENTS" under "DEBT INSTRUMENTS"
        # is matched as a category, not as a new top-level instrument_type

        # Check for Level 2: category_type (children of current Level 1)
        if len(self.current_path) >= 1:
            current_instrument = self.current_path[0]
            children = current_instrument.get('children', [])

            for child in children:
                category_type = child.get('category_type', '')
                normalized_category = self._normalize_cell_text(category_type) if category_type else ''

                if normalized_category and normalized_cell.lower() == normalized_category.lower():
                    # Found a Level 2 marker - close any Level 2/3 if open
                    while len(self.current_path) > 1:
                        self.current_path.pop()
                        self.category_labels.pop()

                    # Push new Level 2
                    self.current_path.append(child)
                    # Normalize category_type to remove extra spaces
                    normalized = normalize_instrument_type(category_type)
                    self.category_labels.append(normalized if normalized is not None else category_type)
                    return True, 'start'

        # Check for Level 3: subcategory_type (children of current Level 2)
        if len(self.current_path) >= 2:
            current_category = self.current_path[1]
            grandchildren = current_category.get('children', [])

            for grandchild in grandchildren:
                subcategory_type = grandchild.get('subcategory_type', '')
                normalized_subcategory = self._normalize_cell_text(subcategory_type) if subcategory_type else ''

                if normalized_subcategory and normalized_cell.lower() == normalized_subcategory.lower():
                    # Found a Level 3 marker - close any Level 3 if open
                    while len(self.current_path) > 2:
                        self.current_path.pop()
                        self.category_labels.pop()

                    # Push new Level 3
                    self.current_path.append(grandchild)
                    # Normalize subcategory_type to remove extra spaces
                    normalized = normalize_instrument_type(subcategory_type)
                    self.category_labels.append(normalized if normalized is not None else subcategory_type)
                    return True, 'start'

        # Check for Level 1: instrument_type (top-level in hierarchy_config)
        # This is checked LAST so that child matches take priority when inside a hierarchy
        for hierarchy_level in self.hierarchy_config:
            instrument_type = hierarchy_level.get('instrument_type', '')
            normalized_instrument = self._normalize_cell_text(instrument_type) if instrument_type else ''

            if normalized_instrument and normalized_cell.lower() == normalized_instrument.lower():
                # Found a Level 1 marker - close all open levels (pop everything)
                self.current_path = []
                self.category_labels = []

                # Push new Level 1
                self.current_path.append(hierarchy_level)
                # Normalize instrument_type to remove extra spaces
                normalized = normalize_instrument_type(instrument_type)
                self.category_labels.append(normalized if normalized is not None else instrument_type)
                return True, 'start'

        return False, None

    def get_current_category(self) -> Optional[str]:
        """Get the current category label."""
        if self.category_labels:
            return ' > '.join(self.category_labels)
        return None

    def get_instrument_type(self) -> Optional[str]:
        """Get the first level (Type of the Instrument)."""
        if len(self.category_labels) >= 1:
            return self.category_labels[0]
        return None

    def get_category(self) -> Optional[str]:
        """Get the second level (Category)."""
        if len(self.category_labels) >= 2:
            return self.category_labels[1]
        return None

    def get_subcategory(self) -> Optional[str]:
        """Get the third level (SubCategory)."""
        if len(self.category_labels) >= 3:
            return self.category_labels[2]
        return None

    def is_in_hierarchy(self) -> bool:
        """Check if we're currently inside a hierarchy."""
        return len(self.current_path) > 0

    def uses_end_keywords(self) -> bool:
        """
        Check if any component in the current hierarchy path uses end_keywords.

        Returns:
            True if any active component has end_keywords defined
        """
        for level in self.current_path:
            if level.get('end_keywords'):
                return True
        return False

    def hierarchy_config_has_end_keywords(self) -> bool:
        """
        Check if ANY component in the entire hierarchy configuration uses end_keywords.

        This is used to determine table-level behavior (whether blank rows should close tables).

        Returns:
            True if any component in the hierarchy config has end_keywords defined
        """
        def check_level(level_config):
            """Recursively check a level and its children for end_keywords."""
            if level_config.get('end_keywords'):
                return True
            for child in level_config.get('children', []):
                if check_level(child):
                    return True
            return False

        for hierarchy_level in self.hierarchy_config:
            if check_level(hierarchy_level):
                return True
        return False

    def should_extract_marker_data(self) -> bool:
        """
        Check if the most recently opened hierarchy level has the extract_data_from_marker flag.

        This is used for special cases like TREPS where the category marker row
        also contains data values that should be extracted.

        Returns:
            True if the current level has extract_data_from_marker set to True
        """
        if not self.current_path:
            return False

        # Check the most recently added level (last in the path)
        current_level = self.current_path[-1]
        return current_level.get('extract_data_from_marker', False)


class TableMetadataLoader:
    """Loads and manages table structure metadata."""

    def __init__(
        self,
        metadata_path: str,
        table_selection_strategy: Optional[str] = None,
        enable_alternate_input_names: Optional[bool] = None,
    ):
        self.metadata_path = metadata_path
        # Table selection strategy used for header matching.
        self.table_selection_strategy = "first_match"

        if enable_alternate_input_names is None:
            env_value = os.getenv("EV_ENABLE_ALTERNATE_INPUT_NAMES", "0").strip().lower()
            self.enable_alternate_input_names = env_value in {"1", "true", "yes", "on"}
        else:
            self.enable_alternate_input_names = bool(enable_alternate_input_names)

        self.table_structures: List[TableStructure] = []
        self.last_match_details: Optional[Dict[str, Any]] = None  # Stores details from last successful match
        self.fund_name_source: Optional[Dict[str, Any]] = None  # Config for extracting fund name from sheet
        self.index_sheet_config: Optional[Dict[str, Any]] = None  # Config for parsing Index sheet
        self.category_overrides: List[Dict[str, Any]] = []  # Content-based category overrides
        self.load_metadata()

    def load_metadata(self):
        """Load table structures from JSON metadata file."""
        try:
            with open(self.metadata_path, 'r') as f:
                metadata = json.load(f)

            # Load fund_name_source config if present
            self.fund_name_source = metadata.get('fund_name_source')

            # Load index_sheet config if present
            self.index_sheet_config = metadata.get('index_sheet')

            # Load category_overrides for content-based category assignment
            self.category_overrides = metadata.get('category_overrides', [])
            if self.category_overrides:
                logger.info(f"Loaded {len(self.category_overrides)} category overrides")

            for table_config in metadata.get('table_structures', []):
                table_structure = TableStructure(table_config)
                table_structure.use_alternate_input_names = self.enable_alternate_input_names
                self.table_structures.append(table_structure)

            # Log metadata dimensionality as requested in specification
            logger.info(f"=== METADATA DIMENSIONALITY ===")
            logger.info(f"Loaded {len(self.table_structures)} table structures from {self.metadata_path}")
            for table_struct in self.table_structures:
                num_columns = len(table_struct.header_columns)
                num_hierarchies = len(table_struct.vertical_hierarchy)
                logger.info(f"  - {table_struct.name}: {num_columns} columns, {num_hierarchies} vertical hierarchies")

        except Exception as e:
            logger.error(f"Failed to load metadata from {self.metadata_path}: {e}")
            raise

    def identify_table_structure(self, row: List[Any], sheet_name: str = None, used_tables: List[str] = None) -> Optional[TableStructure]:
        """
        Identify which table structure matches the given header row.
        Returns the first matching table structure.

        Args:
            row: List of cell values from a row
            sheet_name: Optional sheet name for logging context
            used_tables: List of table names already used in this sheet (for tie-breaking)

        Returns:
            TableStructure if a match is found, None otherwise
        """
        for table_structure in self.table_structures:
            match_details = table_structure.get_header_match_details(row)
            if match_details['is_match']:
                self._apply_selected_match(table_structure, match_details, sheet_name)
                return table_structure

        self.last_match_details = None
        return None

    def _apply_selected_match(
        self,
        table_structure: TableStructure,
        match_details: Dict[str, Any],
        sheet_name: Optional[str],
    ) -> None:
        """Store selected match details and emit partial-match warning when needed."""
        self.last_match_details = match_details
        table_structure.matched_columns_set = set(match_details.get('matched_columns', []))

        if match_details['match_percentage'] < 100:
            sheet_context = f" in sheet '{sheet_name}'" if sheet_name else ""
            logger.warning(
                f"PARTIAL_HEADER_MATCH{sheet_context}: "
                f"{match_details['matched_count']}/{match_details['total_count']} columns "
                f"({match_details['match_percentage']:.1f}%). "
                f"Missing: {match_details['missing_columns']}"
            )

    def apply_category_override(self, record: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Apply category overrides based on content matching.

        Checks if any configured category_override pattern matches the record's
        specified field and updates the category, instrument_type, and/or subcategory accordingly.
        If an override has exclude: true, returns None to indicate the record should be filtered out.

        Args:
            record: The extracted record dictionary

        Returns:
            The record with category/instrument_type/subcategory potentially updated,
            or None if the record should be excluded
        """
        if not self.category_overrides:
            return record

        for override in self.category_overrides:
            pattern = override.get('pattern', '')
            match_field = override.get('match_field', 'instrument_name')
            has_category_override = 'category' in override
            has_instrument_type_override = 'instrument_type' in override
            has_subcategory_override = 'subcategory' in override
            new_category = override.get('category')
            new_instrument_type = override.get('instrument_type')
            new_subcategory = override.get('subcategory')
            exclude = override.get('exclude', False)

            if not pattern:
                continue

            # Check if pattern matches (exclusion or override)
            has_override = has_category_override or has_instrument_type_override or has_subcategory_override
            if not exclude and not has_override:
                continue  # Skip if no exclusion and no override specified

            field_value = record.get(match_field, '')
            if field_value and pattern.lower() in str(field_value).lower():
                # Check for exclusion first
                if exclude:
                    logger.debug(
                        f"Record excluded: '{pattern}' matched in '{match_field}' "
                        f"(instrument_name: '{record.get('instrument_name', 'N/A')}')"
                    )
                    return None  # Signal that this record should be excluded
                
                # Apply overrides (category, instrument_type, subcategory)
                if has_category_override:
                    old_category = record.get('category')
                    record['category'] = normalize_instrument_type(new_category)
                    logger.debug(
                        f"Category override applied: '{pattern}' matched in '{match_field}', "
                        f"changed category from '{old_category}' to '{new_category}'"
                    )
                if has_instrument_type_override:
                    old_instrument_type = record.get('instrument_type')
                    record['instrument_type'] = normalize_instrument_type(new_instrument_type)
                    logger.debug(
                        f"Instrument type override applied: '{pattern}' matched in '{match_field}', "
                        f"changed instrument_type from '{old_instrument_type}' to '{new_instrument_type}'"
                    )
                if has_subcategory_override:
                    old_subcategory = record.get('subcategory')
                    # Normalize subcategory (function is in same module)
                    record['subcategory'] = normalize_instrument_type(new_subcategory)
                    logger.debug(
                        f"Subcategory override applied: '{pattern}' matched in '{match_field}', "
                        f"changed subcategory from '{old_subcategory}' to '{new_subcategory}'"
                    )
                break  # Apply first matching override only

        return record
