"""
Sheet Processor Module

This module provides functionality to process Excel sheets row by row,
identifying tables and extracting data based on table structure metadata.
"""

import pandas as pd
import logging
import re
import warnings
import os
import sys
from contextlib import contextmanager
from typing import List, Dict, Any, Optional, Tuple
from table_structure_parser import TableMetadataLoader, TableStructure, VerticalHierarchyTracker

# Suppress xlrd formula warnings (unknown FuncID warnings)
warnings.filterwarnings('ignore', message='.*formula.*', category=UserWarning)
warnings.filterwarnings('ignore', message='.*tFuncVar.*', category=UserWarning)
warnings.filterwarnings('ignore', message='.*FuncID.*', category=UserWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='xlrd')


@contextmanager
def suppress_stderr():
    """Context manager to suppress stderr output (for xlrd formula warnings)."""
    with open(os.devnull, 'w') as devnull:
        old_stderr = sys.stderr
        try:
            sys.stderr = devnull
            yield
        finally:
            sys.stderr = old_stderr


def should_skip_sheet(name: str, skip_config: Dict[str, List[str]]) -> bool:
    """Check if sheet should be skipped based on config patterns.

    Args:
        name: Sheet name
        skip_config: Dict with 'exact', 'contains', and 'regex' lists

    Returns:
        True if sheet should be skipped
    """
    name_lower = name.lower().strip()
    if name_lower in skip_config.get('exact', []):
        return True
    if any(pattern in name_lower for pattern in skip_config.get('contains', [])):
        return True
    if any(re.match(pattern, name_lower) for pattern in skip_config.get('regex', [])):
        return True
    return False


def _is_header_column_name(text: str) -> bool:
    """Check if text is a header column name (not a fund name).

    Args:
        text: Text to check

    Returns:
        True if text appears to be a header column name
    """
    if not text:
        return False
    text_lower = text.lower().strip()
    return (
        'name of the instrument' in text_lower or
        'name of instrument' in text_lower or
        text_lower.startswith('isin') or
        text_lower.startswith('quantity') or
        text_lower.startswith('market') or
        text_lower.startswith('% to') or
        ('rating' in text_lower and 'industry' in text_lower) or
        'yield' in text_lower or
        'ytm' in text_lower or
        'month ended' in text_lower
    )


def _is_valid_fund_name_candidate(text: str) -> bool:
    """Check if text is a valid fund name candidate.

    Args:
        text: Text to validate

    Returns:
        True if text appears to be a valid fund name
    """
    if not text or len(text.strip()) < 3:
        return False

    text_str = text.strip()
    text_lower = text_str.lower()

    # Reject header column names
    if _is_header_column_name(text_str):
        return False

    # Reject regulatory/disclaimer text
    invalid_keywords = ['pursuant', 'regulation', 'sebi', 'securities and exchange']
    if any(kw in text_lower for kw in invalid_keywords):
        return False

    # Reject scheme descriptions (parentheses with descriptive words) and non-primary label phrases
    non_primary_label_phrases = ['index fund', 'smart beta']
    if (
        (text_str.startswith('(') and any(word in text_lower for word in ['scheme', 'investing', 'securities', 'risk']))
        or any(phrase in text_lower for phrase in non_primary_label_phrases)
    ):
        return False

    # Must contain fund-related keywords (case-insensitive)
    return any(keyword in text_lower for keyword in ['fund', 'portfolio', 'plan', 'scheme', 'fmp', 'etf', 'fof'])


logger = logging.getLogger(__name__)


class SheetProcessor:
    """Processes an Excel sheet row by row to extract structured data."""

    def __init__(self, metadata_loader: TableMetadataLoader):
        self.metadata_loader = metadata_loader
        self.current_table: Optional[TableStructure] = None
        self.hierarchy_tracker: Optional[VerticalHierarchyTracker] = None
        self.records: List[Dict[str, Any]] = []
        self.missing_headers: set = set()  # Track headers in metadata but not in Excel
        self.unrecognized_headers: set = set()  # Track headers in Excel but not in metadata
        self.failed_sheets: set = set()  # Track sheets with data but no extracted records
        self._cached_hierarchy_names: List[str] = []  # Cached hierarchy names for performance

    def process_sheet(self, df: pd.DataFrame, sheet_name: str, fund_name: str, fund_code: str = None) -> List[Dict[str, Any]]:
        """
        Process a single sheet row by row.

        Args:
            df: DataFrame with the sheet data (no headers, raw data)
            sheet_name: Name of the sheet
            fund_name: Full name of the fund
            fund_code: Fund code (defaults to sheet_name if not provided)

        Returns:
            List of extracted records
        """
        self.records = []
        self.current_table = None
        self.hierarchy_tracker = None
        self._cached_hierarchy_names = []
        self._seen_net_receivables = set()  # Track which net receivables texts have been seen
        self._used_tables = []  # Track table names used in this sheet (for sequential identical-header tables)
        self._df = df  # Store DataFrame reference for lookahead checks
        # Use provided fund_code (can be None if explicitly blank) or fall back to sheet_name
        # Note: fund_code can be intentionally None (e.g., when AMC doesn't provide codes)
        self._fund_code = fund_code

        logger.info(f"Processing sheet: {sheet_name} (fund_code: {self._fund_code}, fund_name: {fund_name})")

        # Check if sheet has any populated cells
        has_populated_cells = False
        for idx in range(len(df)):
            row = df.iloc[idx].tolist()
            if any(cell is not None and not (isinstance(cell, float) and str(cell) == 'nan') and str(cell).strip() for cell in row):
                has_populated_cells = True
                break

        # Iterate through each row
        for idx in range(len(df)):
            row = df.iloc[idx].tolist()

            # Check if row is completely empty
            is_empty = not any(
                cell is not None and
                not (isinstance(cell, float) and str(cell) == 'nan') and
                str(cell).strip()
                for cell in row
            )

            if is_empty:
                # Check if table or hierarchy configuration uses end_keywords
                # If yes: blank rows should NOT close the table (only end_keywords can close components)
                # If no: blank rows close the table (original behavior)
                # BUT: If table has a hierarchy structure, don't close on first empty row after header
                # (there might be spacing before hierarchy markers start)
                if self.current_table and self.hierarchy_tracker:
                    # Check both table-level and hierarchy-level end_keywords
                    has_table_end_keywords = bool(self.current_table.table_end_keywords)
                    has_hierarchy_end_keywords = self.hierarchy_tracker.hierarchy_config_has_end_keywords()
                    has_hierarchy = len(self.current_table.vertical_hierarchy) > 0
                    
                    if has_table_end_keywords or has_hierarchy_end_keywords:
                        # Table or hierarchy uses end_keywords - ignore blank rows within the table
                        logger.debug(f"Row {idx + 1}: Empty row in end_keyword mode, continuing...")
                        continue
                    elif has_hierarchy:
                        # Table has hierarchy - be lenient with empty rows (spacing between hierarchy markers is common)
                        # Only close table if we've seen multiple consecutive empty rows (likely end of table)
                        # For now, allow empty rows within hierarchy (they're just spacing)
                        logger.debug(f"Row {idx + 1}: Empty row in hierarchy, allowing spacing...")
                        continue
                    else:
                        # No end_keywords and no hierarchy - blank row closes the table (original behavior)
                        logger.debug(f"Row {idx + 1}: Empty row detected, ending current table '{self.current_table.name}'")
                        self._used_tables.append(self.current_table.name)
                        self.current_table = None
                        self.hierarchy_tracker = None
                        continue
                # If no active table, just skip the empty row
                continue

            # Try to identify if this row is a table header
            table_structure = self.metadata_loader.identify_table_structure(row, sheet_name=sheet_name, used_tables=self._used_tables)

            if table_structure:
                # Found a new table header
                logger.debug(f"Found table '{table_structure.name}' at row {idx + 1}")
                self.current_table = table_structure
                self.hierarchy_tracker = VerticalHierarchyTracker(
                    table_structure.vertical_hierarchy,
                    hierarchy_marker_column=table_structure.hierarchy_marker_column,
                    default_instrument_type=table_structure.default_instrument_type
                )
                # Cache hierarchy names for performance
                self._cached_hierarchy_names = self._get_hierarchy_names(table_structure.vertical_hierarchy)
                # Reset seen net receivables for new table
                self._seen_net_receivables = set()

                # For multi-portfolio sheets: look for fund/portfolio name in rows above header
                # Only update if this is NOT the first table (idx > 5 means we're past the initial header area)
                # AND fund_name was not extracted via fund_name_source (i.e., it's still the default sheet_name)
                # This prevents overriding a valid fund_name that was extracted from metadata
                if idx > 5 and fund_name == sheet_name:
                    candidate_fund_name = None
                    for lookback in range(1, min(6, idx + 1)):
                        prev_row = df.iloc[idx - lookback].tolist()
                        first_cell = prev_row[0] if len(prev_row) > 0 else None
                        second_cell = prev_row[1] if len(prev_row) > 1 else None

                        # Detect if first_cell is a prefix/number (like "1)", "2)") - if so, use column 1 instead
                        first_is_prefix = False
                        if first_cell is not None and pd.notna(first_cell):
                            first_str = str(first_cell).strip()
                            first_is_prefix = (
                                (first_str.endswith(')') and len(first_str) <= 5) or
                                (len(first_str) <= 3 and first_str.replace('.', '').isdigit())
                            )

                        # Use second_cell if first is a prefix, otherwise use first
                        cell_to_check = second_cell if first_is_prefix and second_cell is not None and pd.notna(second_cell) else first_cell

                        if cell_to_check is not None and pd.notna(cell_to_check):
                            cell_str = str(cell_to_check).strip()

                            # Skip header columns
                            if _is_header_column_name(cell_str):
                                continue

                            # Check if it's a valid fund name candidate (length > 15 for substantial names)
                            if len(cell_str) > 15 and _is_valid_fund_name_candidate(cell_str):
                                # Skip if it's clearly a date/statement row, notes, or disclaimer
                                cell_lower = cell_str.lower()
                                is_disclaimer = (
                                    cell_str.startswith('Portfolio Statement') or
                                    cell_str.startswith('Portfolio as on') or
                                    cell_str.startswith('Notes') or
                                    cell_str.startswith('None of the schemes') or
                                    cell_str.startswith('$$$') or
                                    cell_str.startswith('**') or
                                    cell_str.startswith('Post the last') or
                                    cell_str.startswith('As on ') or
                                    'NAV at the' in cell_str or
                                    'winding-up' in cell_lower or
                                    'liquidator' in cell_lower or
                                    'investors may note' in cell_lower or
                                    'stand extinguished' in cell_lower or
                                    'riskometer' in cell_lower or
                                    '100% of' in cell_lower or
                                    'paid out to investors' in cell_lower or
                                    'paid 100%' in cell_lower or
                                    'aum' in cell_lower
                                )
                                if not is_disclaimer:
                                    # Prefer longer names (more complete fund names)
                                    if candidate_fund_name is None or len(cell_str) > len(candidate_fund_name):
                                        candidate_fund_name = cell_str
                    if candidate_fund_name:
                        fund_name = candidate_fund_name
                        logger.debug(f"Multi-portfolio: Updated fund_name: {fund_name[:60]}...")

                # Track missing and unrecognized headers if partial match
                match_details = self.metadata_loader.last_match_details
                if match_details:
                    if match_details.get('missing_columns'):
                        self.missing_headers.update(match_details['missing_columns'])
                    if match_details.get('unrecognized_columns'):
                        self.unrecognized_headers.update(match_details['unrecognized_columns'])

                continue

            # If we have an active table, process the row
            if self.current_table and self.hierarchy_tracker:
                self._process_data_row(row, sheet_name, fund_name, idx + 1, df, idx)

        # Failsafe: If sheet has populated cells but no records were extracted, track it
        # Ignore 'Index' sheet (case insensitive)
        if has_populated_cells and len(self.records) == 0 and sheet_name.lower() != 'index':
            self.failed_sheets.add(sheet_name)
            logger.error(f"FAILSAFE: Sheet '{sheet_name}' has populated cells but NO records were extracted!")

        logger.info(f"Extracted {len(self.records)} records from sheet {sheet_name}")
        return self.records

    def _process_data_row(self, row: List[Any], sheet_name: str, fund_name: str, row_num: int, df: pd.DataFrame = None, row_idx: int = None):
        """
        Process a data row within an active table.

        Args:
            row: List of cell values
            sheet_name: Sheet name (fund code)
            fund_name: Fund name
            row_num: Row number (1-indexed)
            df: DataFrame reference for lookahead checks (optional)
            row_idx: Row index in DataFrame for lookahead checks (optional)
        """
        # Check for table_end_keywords - if found, close the table
        # But first check if this is also a Net Receivables row (extract it before closing)
        # net_receivables_match is the matched text or empty string if no match
        net_receivables_match = self._is_net_receivables_row(row)

        if self.current_table.table_end_keywords:
            # Get text to check for table end keywords
            # If table_end_check_all_cells is True, check ALL cells (for footer text in middle columns)
            # Otherwise, check only the first non-empty cell (default behavior)
            if self.current_table.table_end_check_all_cells:
                # Join all non-empty cells into a single string
                check_text = ' '.join([
                    str(cell).strip() for cell in row
                    if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan') and str(cell).strip()
                ])
            else:
                # Default: get first non-empty cell value only
                check_text = None
                for cell in row:
                    if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan'):
                        cell_str = str(cell).strip()
                        if cell_str:
                            check_text = cell_str
                            break

            if check_text:
                for keyword in self.current_table.table_end_keywords:
                    # Support prefix matching: keywords starting with "^" match start of check_text
                    keyword_lower = keyword.lower()
                    check_text_lower = check_text.lower()
                    if keyword_lower.startswith('^'):
                        matches = check_text_lower.startswith(keyword_lower[1:])
                    else:
                        matches = keyword_lower in check_text_lower

                    # Special handling for "TOTAL" - check if it's followed by exclude patterns (footnotes)
                    # This handles segregated portfolios that end with TOTAL instead of GRAND TOTAL
                    if matches and keyword_lower == 'total' and 'grand' not in check_text_lower:
                        # Check if this TOTAL is followed by exclude patterns (indicating final table end)
                        # vs more sections (indicating intermediate section total)
                        if (self.current_table.close_table_after_total_on_exclude_match and
                            df is not None and row_idx is not None):
                            is_final_total = self._check_total_followed_by_exclude_patterns(df, row_idx)
                            if not is_final_total:
                                # This is an intermediate TOTAL, don't close the table
                                continue

                    if matches:
                        # If this is also a Net Receivables row, extract it first
                        if net_receivables_match:
                            record = self.current_table.extract_record(row)
                            # Set instrument_name BEFORE validation so alphabetic check passes
                            record['instrument_name'] = net_receivables_match
                            if self._is_valid_record(record):
                                record['fund_code'] = self._fund_code
                                record['fund_name'] = fund_name
                                record['table_type'] = self.current_table.name
                                # Use net_receivables overrides if set (empty string "" means blank)
                                if self.current_table.net_receivables_instrument_type is not None:
                                    record['instrument_type'] = self.current_table.net_receivables_instrument_type or None
                                else:
                                    record['instrument_type'] = self.hierarchy_tracker.get_instrument_type()
                                if self.current_table.net_receivables_category is not None:
                                    record['category'] = self.current_table.net_receivables_category or None
                                else:
                                    record['category'] = self.hierarchy_tracker.get_category()
                                record['subcategory'] = self.hierarchy_tracker.get_subcategory()
                                self._apply_exclude_columns(record)
                                self.records.append(record)
                                logger.debug(f"Row {row_num}: Extracted Net Receivables record before closing table")

                        logger.debug(f"Row {row_num}: Table end keyword '{keyword}' found, closing table '{self.current_table.name}'")
                        self._used_tables.append(self.current_table.name)
                        self.current_table = None
                        self.hierarchy_tracker = None
                        return

        # Check if this is a Net Receivables row (special case - no hierarchy needed)
        # Track seen receivables and close table only after ALL expected ones are seen
        if net_receivables_match:
            record = self.current_table.extract_record(row)
            # Set instrument_name BEFORE validation so alphabetic check passes
            record['instrument_name'] = net_receivables_match
            # Only extract if row has meaningful data (market_value, percent_to_nav, etc.)
            # This handles cases where net_receivables_text appears as a hierarchy header row
            # followed by a data row with the same text but actual values
            if self._has_meaningful_data(record) and self._is_valid_record(record):
                # Clear ISIN field - net receivables rows don't have ISINs
                # (the net_receivables_text may appear in the ISIN column position in some formats)
                record['isin'] = None

                # Add metadata with hierarchy fields from current tracker state
                record['fund_code'] = self._fund_code
                record['fund_name'] = fund_name
                record['table_type'] = self.current_table.name
                # Use net_receivables overrides if set (empty string "" means blank)
                if self.current_table.net_receivables_instrument_type is not None:
                    record['instrument_type'] = self.current_table.net_receivables_instrument_type or None
                else:
                    record['instrument_type'] = self.hierarchy_tracker.get_instrument_type()
                if self.current_table.net_receivables_category is not None:
                    record['category'] = self.current_table.net_receivables_category or None
                else:
                    record['category'] = self.hierarchy_tracker.get_category()
                record['subcategory'] = self.hierarchy_tracker.get_subcategory()
                self._apply_exclude_columns(record)
                self.records.append(record)
                logger.debug(f"Row {row_num}: Extracted Net Receivables record '{net_receivables_match}'")

                # Track this net receivables text as seen (only when successfully extracted)
                self._seen_net_receivables.add(net_receivables_match.lower())

                # Check if ALL expected net receivables have been seen
                expected_texts = set(t.lower() for t in self.current_table.net_receivables_texts)
                if self._seen_net_receivables >= expected_texts:
                    # All net receivables seen - close the table
                    logger.debug(f"Row {row_num}: All net receivables seen, closing table '{self.current_table.name}'")
                    self._used_tables.append(self.current_table.name)
                    self.current_table = None
                    self.hierarchy_tracker = None
                return
            # If no meaningful data, don't return - let this row be checked as a hierarchy marker

        # Check if this row is a hierarchy marker
        # Pass data column indices so the tracker can distinguish data rows from markers
        data_column_indices = [col_def['index'] for col_def in self.current_table.header_columns
                               if col_def['output_name'] in ['isin', 'quantity', 'market_value', 'percent_to_nav', 'percent_to_aum', 'ytm', 'ytc']]
        # Get instrument_name column index for checking hierarchy markers
        instrument_name_col_idx = next((col_def['index'] for col_def in self.current_table.header_columns
                                       if col_def['output_name'] == 'instrument_name'), None)
        is_marker, marker_type = self.hierarchy_tracker.check_and_update(row, data_column_indices, instrument_name_col_idx)

        if is_marker:
            # For 'start' markers only: check if row has meaningful data
            # This handles cases where instrument_name matches a category_type but row contains actual data
            # NOTE: 'end' markers (Sub Total, Total, Grand Total) should NEVER be extracted as data
            if marker_type == 'start':
                record = self.current_table.extract_record(row)
                if self._has_meaningful_data(record) and self._is_valid_record(record):
                    # This row has meaningful data - treat as data row, not pure hierarchy marker
                    record['fund_code'] = self._fund_code
                    record['fund_name'] = fund_name
                    record['table_type'] = self.current_table.name
                    record['instrument_type'] = (
                        self.current_table.default_instrument_type or
                        self.hierarchy_tracker.get_instrument_type()
                    )
                    record['category'] = self.hierarchy_tracker.get_category()
                    record['subcategory'] = self.hierarchy_tracker.get_subcategory()
                    # Check if instrument should have blank hierarchy
                    instrument_name = record.get('instrument_name', '')
                    if instrument_name and self.current_table.blank_hierarchy_instruments:
                        if instrument_name.lower().strip() in self.current_table.blank_hierarchy_instruments:
                            record['instrument_type'] = None
                            record['category'] = None
                    self._apply_exclude_columns(record)
                    self.records.append(record)
                    logger.debug(f"Row {row_num}: Hierarchy marker ({marker_type}) but has meaningful data - extracted record")
                    return

            # Check if this marker should also be extracted as data (legacy flag)
            if marker_type == 'start' and self.hierarchy_tracker.should_extract_marker_data():
                record = self.current_table.extract_record(row)
                # This marker is configured to have its data extracted
                if self._is_valid_record(record):
                    record['fund_code'] = self._fund_code
                    record['fund_name'] = fund_name
                    record['table_type'] = self.current_table.name
                    # Normalize instrument_type (already normalized at source, but ensure consistency)
                    from table_structure_parser import normalize_instrument_type
                    instrument_type = (
                        self.current_table.default_instrument_type or
                        self.hierarchy_tracker.get_instrument_type()
                    )
                    record['instrument_type'] = normalize_instrument_type(instrument_type)
                    record['category'] = self.hierarchy_tracker.get_category()
                    record['subcategory'] = self.hierarchy_tracker.get_subcategory()
                    # Check if instrument should have blank hierarchy
                    instrument_name = record.get('instrument_name', '')
                    if instrument_name and self.current_table.blank_hierarchy_instruments:
                        if instrument_name.lower().strip() in self.current_table.blank_hierarchy_instruments:
                            record['instrument_type'] = None
                            record['category'] = None
                    self._apply_exclude_columns(record)
                    self.records.append(record)
                    logger.debug(f"Row {row_num}: Hierarchy marker ({marker_type}) with data extraction flag - extracted record")
                else:
                    logger.debug(f"Row {row_num}: Hierarchy marker ({marker_type}) with data extraction flag, but no valid data")
            else:
                # Normal hierarchy marker without data extraction
                logger.debug(f"Row {row_num}: Hierarchy marker ({marker_type})")
            return

        # Check if we're inside a hierarchy (ready to extract data)
        # If table has no hierarchy (empty vertical_hierarchy), allow all data rows
        has_hierarchy = len(self.current_table.vertical_hierarchy) > 0
        if has_hierarchy and not self.hierarchy_tracker.is_in_hierarchy():
            # We're not inside any hierarchy block, skip this row
            return

        # Extract the record
        record = self.current_table.extract_record(row)

        # Filter out invalid records
        if not self._is_valid_record(record):
            return

        # Add metadata
        record['fund_code'] = self._fund_code
        record['fund_name'] = fund_name
        record['table_type'] = self.current_table.name
        # Use default_instrument_type if set, otherwise use hierarchy
        # Note: Both are already normalized at source, but normalize here as safety measure
        instrument_type = (
            self.current_table.default_instrument_type or
            self.hierarchy_tracker.get_instrument_type()
        )
        # Import normalize function
        from table_structure_parser import normalize_instrument_type
        record['instrument_type'] = normalize_instrument_type(instrument_type)
        # Normalize category and subcategory for consistency (already normalized at source)
        record['category'] = normalize_instrument_type(self.hierarchy_tracker.get_category())
        record['subcategory'] = normalize_instrument_type(self.hierarchy_tracker.get_subcategory())

        # Check if instrument should have blank hierarchy (instrument_type and category)
        instrument_name = record.get('instrument_name', '')
        if instrument_name and self.current_table.blank_hierarchy_instruments:
            if instrument_name.lower().strip() in self.current_table.blank_hierarchy_instruments:
                record['instrument_type'] = None
                record['category'] = None
                logger.debug(f"Row {row_num}: Blanked hierarchy for '{instrument_name}' (blank_hierarchy_instruments)")

        # Apply category overrides based on content matching
        record = self.metadata_loader.apply_category_override(record)
        
        # If record was excluded by category override, skip it
        if record is None:
            return

        # Remove excluded output columns (if configured)
        # NOTE: Do this AFTER category_overrides, as overrides may need these fields for matching
        self._apply_exclude_columns(record)

        self.records.append(record)
        logger.debug(f"Row {row_num}: Extracted record for '{record.get('instrument_name', 'N/A')}'")

    def _apply_exclude_columns(self, record: Dict[str, Any]) -> None:
        """Remove excluded output columns from a record.

        Args:
            record: The record dict to modify in place
        """
        if self.current_table and self.current_table.exclude_output_columns:
            for col_name in self.current_table.exclude_output_columns:
                record.pop(col_name, None)

    def _is_net_receivables_row(self, row: List[Any]) -> str:
        """
        Check if a row is a Net Receivables row.

        Args:
            row: List of cell values from a row

        Returns:
            The matched net receivables text if found, empty string otherwise
        """
        # Check if current table has net_receivables_texts defined
        if not self.current_table or not self.current_table.net_receivables_texts:
            return ""

        # Convert row to string for checking
        row_str = ' '.join([str(cell).lower() for cell in row
                            if cell is not None and
                            not (isinstance(cell, float) and str(cell) == 'nan') and
                            str(cell).strip()])

        # Check if any net_receivables_text is contained in the row
        for net_recv_text in self.current_table.net_receivables_texts:
            if net_recv_text.lower() in row_str:
                return net_recv_text
        return ""

    def _get_hierarchy_names(self, hierarchy_config: List[Dict]) -> List[str]:
        """Recursively extract all hierarchy marker names."""
        names = []
        for level in hierarchy_config:
            if 'instrument_type' in level:
                names.append(level['instrument_type'].lower().strip())
            if 'category_type' in level:
                names.append(level['category_type'].lower().strip())
            if 'subcategory_type' in level:
                names.append(level['subcategory_type'].lower().strip())
            if 'children' in level:
                names.extend(self._get_hierarchy_names(level['children']))
        return names

    def _check_total_followed_by_exclude_patterns(self, df: pd.DataFrame, row_idx: int) -> bool:
        """
        Check if a TOTAL row is followed by exclude patterns (footnotes) vs more sections.
        
        Returns True if TOTAL is followed by exclude patterns (final table end).
        Returns False if TOTAL is followed by more sections (intermediate total).
        
        This is scalable across AMCs by:
        1. Using metadata-defined data column indices (not hardcoded)
        2. Deriving section keywords from vertical_hierarchy (not hardcoded)
        3. Checking all columns for exclude patterns (not just column 1)
        """
        if not self.current_table or not self.current_table.exclude_instrument_patterns:
            return False
        
        # Get data column indices from metadata (scalable across AMCs)
        data_column_indices = [
            col_def['index'] for col_def in self.current_table.header_columns 
            if col_def['output_name'] in ['quantity', 'market_value', 'percent_to_nav', 'percent_to_aum', 'yield', 'ytm', 'ytc']
        ]
        
        # Derive section keywords from vertical_hierarchy (scalable across AMCs)
        section_keywords = self._get_hierarchy_names(self.current_table.vertical_hierarchy)
        # Also check for common table end keywords
        section_keywords.extend([kw.lower() for kw in self.current_table.table_end_keywords])
        
        # Look ahead up to 5 rows
        for lookahead_idx in range(row_idx + 1, min(len(df), row_idx + 6)):
            lookahead_row = df.iloc[lookahead_idx].tolist()
            
            # Check ALL columns for exclude patterns (not just column 1)
            lookahead_text = ' '.join([
                str(cell).strip() for cell in lookahead_row
                if cell is not None and not (isinstance(cell, float) and str(cell) == 'nan') and str(cell).strip()
            ]).lower()
            
            # Check if row matches exclude patterns
            matches_exclude = any(
                pattern.lower() in lookahead_text 
                for pattern in self.current_table.exclude_instrument_patterns
            )
            
            # Check if row has data columns (using metadata-defined indices)
            has_data = False
            if data_column_indices:
                has_data = any(
                    idx < len(lookahead_row) and
                    isinstance(lookahead_row[idx], (int, float)) and 
                    not pd.isna(lookahead_row[idx]) and 
                    lookahead_row[idx] != 0
                    for idx in data_column_indices
                )
            else:
                # Fallback: check if any numeric value exists (generic check)
                has_data = any(
                    isinstance(cell, (int, float)) and not pd.isna(cell) and cell != 0
                    for cell in lookahead_row[2:]  # Skip first 2 columns (usually text)
                )
            
            # If matches exclude patterns and has no data -> footnotes (final TOTAL)
            if matches_exclude and not has_data:
                return True
            
            # If has more sections (from hierarchy) -> intermediate TOTAL
            if any(section.lower() in lookahead_text for section in section_keywords):
                return False
        
        # If we reach here, no clear indicator - default to False (don't close)
        return False

    def _has_meaningful_data(self, record: Dict[str, Any]) -> bool:
        """Check if record has numeric data indicating it's a real holding."""
        def is_nonzero(val):
            """Check if value is non-None and non-zero (handles strings)."""
            if val is None:
                return False
            try:
                return float(val) != 0
            except (ValueError, TypeError):
                return False

        has_market_value = is_nonzero(record.get('market_value'))
        has_percent = is_nonzero(record.get('percent_to_nav')) or is_nonzero(record.get('percent_to_aum'))
        has_quantity = is_nonzero(record.get('quantity'))
        has_ytm = is_nonzero(record.get('ytm')) or is_nonzero(record.get('ytm_percent'))
        return has_market_value or has_percent or has_quantity or has_ytm

    def _is_valid_record(self, record: Dict[str, Any]) -> bool:
        """
        Check if a record is valid and should be included.

        Per specification line 36: "If a row has the string 'total' in the instrument name
        column, and there are only numerical (including decimals and percentages) values in
        the other columns, that row should be filtered out."

        Args:
            record: Extracted record dictionary

        Returns:
            True if the record is valid
        """
        # Count populated columns with meaningful data
        # A column is "populated" if it has a non-None, non-zero, non-empty value
        populated_columns = 0
        for value in record.values():
            if value is None:
                continue
            # Check if value is meaningful (not zero, not empty string)
            value_str = str(value).strip()
            if value_str and value_str.lower() != 'nan':
                # Check if it's a zero value
                try:
                    if float(value_str) == 0.0:
                        continue  # Zero values don't count as populated
                except (ValueError, TypeError):
                    pass  # Not a number, so it's meaningful
                populated_columns += 1

        # If completely empty or only zeros, skip
        if populated_columns == 0:
            return False

        # Filter out records where all data fields are "NIL", "N/A", or similar placeholders
        metadata_fields = ['fund_code', 'fund_name', 'table_type', 'instrument_type',
                          'category', 'subcategory', 'instrument_name', 'issuer_name', 'industry']
        data_fields = {k: v for k, v in record.items() if k not in metadata_fields and v is not None}

        if data_fields:
            all_placeholders = True
            for value in data_fields.values():
                value_str = str(value).strip().upper()
                if value_str not in ['NIL', 'N/A', 'NA', '-', '']:
                    all_placeholders = False
                    break

            if all_placeholders:
                return False

        # Filter out rows with "total" in instrument name (Sub Total, Total, GRAND TOTAL, etc.)
        instrument_name = record.get('instrument_name', '')

        # Filter out records with empty or missing instrument_name
        if not instrument_name or (isinstance(instrument_name, str) and not instrument_name.strip()):
            return False

        # Skip rows where no field has alphabetic characters
        # (these are typically subtotal rows with only numeric values)
        # "nan" strings are not counted as having alphabetic characters
        has_alpha = False
        for value in record.values():
            if value is None:
                continue
            value_str = str(value).strip().lower()
            if value_str == 'nan':
                continue
            if any(c.isalpha() for c in value_str):
                has_alpha = True
                break
        if not has_alpha:
            return False

        if isinstance(instrument_name, str):
            instrument_lower = instrument_name.lower().strip()

            # Exact match filter for subtotal/total rows (these should never be data records)
            # regardless of whether they have other data like ISIN
            if instrument_lower in ('sub total', 'subtotal', 'sub-total', 'total', 'grand total'):
                return False

            # Filter out records where instrument_name is just punctuation or very short meaningless text
            # e.g., ',', '.', '-', etc.
            import re
            if re.match(r'^[\W_]+$', instrument_lower) or len(instrument_lower) < 2:
                return False

            # Filter out hierarchy marker names that shouldn't be captured as data
            # Check against cached hierarchy markers from the current table's metadata
            # BUT only filter if they don't have meaningful numeric data (to avoid filtering actual holdings)
            if self._cached_hierarchy_names:
                # Check if this row has meaningful numeric data - if so, it's real data, not a label
                has_meaningful_data = self._has_meaningful_data(record)

                # Normalize instrument_name and check for exact match (allowing whitespace variations)
                instrument_normalized = ' '.join(instrument_lower.split())

                # Also strip prefixes like "(A)", "(I)", "(II)", "A)", "i)" for hierarchy matching
                # These prefixes are used for numbering sections but the base name should match
                instrument_prefix_stripped = re.sub(r'^\([a-z]+\)\s*', '', instrument_normalized, flags=re.IGNORECASE)
                instrument_prefix_stripped = re.sub(r'^[a-z]\)\s*', '', instrument_prefix_stripped, flags=re.IGNORECASE)
                instrument_prefix_stripped = re.sub(r'^[ivxlcdm]+\)\s*', '', instrument_prefix_stripped, flags=re.IGNORECASE)

                for hierarchy_name in self._cached_hierarchy_names:
                    hierarchy_normalized = ' '.join(hierarchy_name.split())
                    # Check both original and prefix-stripped versions
                    if instrument_normalized == hierarchy_normalized or instrument_prefix_stripped == hierarchy_normalized:
                        # This is an exact match to a hierarchy marker
                        # Only filter if it doesn't have meaningful data (it's just a label, not a holding)
                        if not has_meaningful_data:
                            return False
                        # If it has data, it's a real holding that happens to match a hierarchy name
                        break

                    # Also check if instrument_name is a partial match (e.g., "TREPS" matches "TREPS / Reverse REPO")
                    # Split hierarchy name and check if instrument_name matches any part
                    hierarchy_parts = [part.strip() for part in hierarchy_normalized.split('/')]
                    for part in hierarchy_parts:
                        if part and (instrument_normalized == part or instrument_prefix_stripped == part):
                            # Instrument name matches a part of a hierarchy marker
                            # Only filter if it doesn't have meaningful data (it's just a label, not a holding)
                            if not has_meaningful_data:
                                return False
                            # If it has data, it's a real holding that happens to match a hierarchy part
                            break

            # Filter out URLs (footnote links that get captured as instrument names)
            if instrument_lower.startswith('http://') or instrument_lower.startswith('https://'):
                return False

            # Filter out known DSP footnote text that can be captured as instrument rows
            if 'net assets does not include' in instrument_lower:
                return False

            # Filter out section headers/notes that appear in tables (global defaults)
            if ('securities classified as below investment grade' in instrument_lower or
                'non traded securities/illiquid securities' in instrument_lower):
                return False

            # Filter out patterns defined in metadata (AMC-specific footnotes/disclaimers)
            if self.current_table and self.current_table.exclude_instrument_patterns:
                for pattern in self.current_table.exclude_instrument_patterns:
                    if pattern.lower() in instrument_lower:
                        return False

            if 'total' in instrument_lower:
                # Check if other columns are only numerical values
                # Skip fund_code, fund_name, table_type, instrument_type, category, subcategory
                # and instrument_name itself when checking for numerical values
                metadata_fields = ['fund_code', 'fund_name', 'table_type', 'instrument_type',
                                   'category', 'subcategory', 'instrument_name', 'issuer_name', 'industry']
                data_fields = {k: v for k, v in record.items() if k not in metadata_fields and v is not None}

                # If all remaining fields are numerical (or can be converted to numerical), filter out
                all_numerical = True
                for value in data_fields.values():
                    if isinstance(value, (int, float)):
                        continue
                    # Treat "#" as numerical (it means <0.005% in these files)
                    value_str = str(value).replace('%', '').replace(',', '').strip()
                    if value_str == '#':
                        continue  # Treat as numerical placeholder
                    try:
                        float(value_str)
                    except (ValueError, AttributeError):
                        all_numerical = False
                        break

                if all_numerical:
                    return False

        # Filter out rows where "total" appears in industry/rating column and instrument_name is empty
        # These are subtotal rows that should not be extracted
        industry_rating = record.get('industry/rating', '')
        instrument_name = record.get('instrument_name', '')
        if industry_rating and isinstance(industry_rating, str):
            if 'total' in industry_rating.lower():
                # If instrument_name is empty/None and we have "total" in industry/rating, it's a subtotal row
                if not instrument_name or (isinstance(instrument_name, str) and not instrument_name.strip()):
                    return False

        # Filter out rows where "total" appears in isin column (Sub Total, Total, Grand Total rows)
        # These are summary rows that have values in market_value/percent columns but no actual instrument
        isin = record.get('isin', '')
        if isin and isinstance(isin, str):
            isin_lower = isin.lower().strip()
            if 'total' in isin_lower or isin_lower == 'grand total' or isin_lower == 'sub total':
                return False

        return True


class ExcelFileProcessor:
    """Processes an entire Excel file with multiple sheets."""

    # Default skip sheets config
    DEFAULT_SKIP_SHEETS = {
        'exact': ['index'],
        'contains': ['riskometer'],
        'regex': []
    }

    def __init__(
        self,
        excel_path: str,
        metadata_path: str,
        skip_sheets: Dict[str, List[str]] = None,
        table_selection_strategy: Optional[str] = None,
        enable_alternate_input_names: Optional[bool] = None,
    ):
        self.excel_path = excel_path
        self.metadata_loader = TableMetadataLoader(
            metadata_path,
            table_selection_strategy=table_selection_strategy,
            enable_alternate_input_names=enable_alternate_input_names,
        )
        self.fund_index: Dict[str, str] = {}
        self.all_records: List[Dict[str, Any]] = []
        self.missing_headers: set = set()  # Track headers in metadata but not in Excel
        self.unrecognized_headers: set = set()  # Track headers in Excel but not in metadata
        self.failed_sheets: set = set()  # Track sheets with data but no extracted records

        # Normalize skip_sheets config (lowercase all patterns)
        raw_config = skip_sheets or self.DEFAULT_SKIP_SHEETS
        self.skip_sheets = {
            'exact': [s.lower() for s in raw_config.get('exact', [])],
            'contains': [s.lower() for s in raw_config.get('contains', [])],
            'regex': raw_config.get('regex', [])
        }

    def parse_index_sheet(self, excel_file: pd.ExcelFile) -> Tuple[Dict[str, str], List[Tuple[str, str]]]:
        """
        Parse the Index sheet to get fund code to fund name mapping.

        Tries to find index sheet by (in order of priority):
        1. Exact match for configured index_sheet.sheet_name (from metadata) - allows per-AMC override
        2. Exact match for 'Index' - most common case, fastest path
        3. Pattern matching for sheets containing 'index' (case-insensitive) - handles variations like "Index -fornight"

        This approach is:
        - Backward compatible: existing AMCs with 'Index' sheets continue to work
        - Scalable: new AMCs can configure sheet_name in metadata if needed
        - Efficient: checks most common case ('Index') before doing pattern matching
        - Future-proof: handles edge cases like typos in sheet names

        Args:
            excel_file: Opened Excel file

        Returns:
            Tuple of:
                - Dictionary mapping fund codes to fund names
                - Ordered list of (fund_code, fund_name) tuples in the order they appear
        """
        # Get index_sheet config from metadata if available
        index_config = self.metadata_loader.index_sheet_config
        configured_sheet_name = index_config.get('sheet_name') if index_config else None
        
        # Get sheet names once (efficient - only called once per file)
        sheet_names = excel_file.sheet_names
        
        # Find the index sheet using priority order
        index_sheet_name = None
        
        # Priority 1: Use configured sheet name if specified (allows per-AMC override)
        if configured_sheet_name:
            if configured_sheet_name in sheet_names:
                index_sheet_name = configured_sheet_name
                logger.debug(f"Using configured index sheet: '{index_sheet_name}'")
            else:
                logger.warning(f"Configured index sheet '{configured_sheet_name}' not found in file")
        
        # Priority 2: Try exact 'Index' match (most common case, fastest)
        if not index_sheet_name and 'Index' in sheet_names:
            index_sheet_name = 'Index'
        
        # Priority 3: Pattern matching for variations (only if above didn't match)
        if not index_sheet_name:
            # Find all sheets containing 'index' (case-insensitive)
            index_candidates = [s for s in sheet_names if 'index' in s.lower()]
            if index_candidates:
                # Prefer shortest match (most likely to be correct, e.g., "Index" > "Index -fornight")
                # If multiple matches, shortest is usually the intended one
                index_sheet_name = min(index_candidates, key=len)
                logger.debug(f"Using pattern-matched index sheet: '{index_sheet_name}' (from {len(index_candidates)} candidates)")
        
        if not index_sheet_name:
            logger.warning("No index sheet found in Excel file")
            return {}, []
        
        try:
            # Suppress stderr for xlrd formula warnings
            with suppress_stderr():
                df = pd.read_excel(excel_file, sheet_name=index_sheet_name, header=None)

            fund_mapping = {}
            fund_list_ordered = []  # Preserve order for positional mapping

            # Get index_sheet config from metadata (if available)
            # Default: column 1 = fund_code, column 2 = fund_name (standard format)
            fund_code_col = index_config.get('fund_code_col', 1) if index_config else 1
            fund_name_col = index_config.get('fund_name_col', 2) if index_config else 2

            # Look for rows with fund code and fund name
            for idx in range(len(df)):
                row = df.iloc[idx]

                # Skip rows that don't have enough columns
                max_col_needed = max(fund_code_col, fund_name_col) + 1
                if len(row) < max_col_needed:
                    continue

                # Get fund code and fund name from configured columns
                fund_code = row.iloc[fund_code_col] if len(row) > fund_code_col else None
                fund_name = row.iloc[fund_name_col] if len(row) > fund_name_col else None

                # Skip if either value is null/NaN
                if not pd.notna(fund_code) or not pd.notna(fund_name):
                    continue

                # Convert to string and strip whitespace
                code = str(fund_code).strip()
                name = str(fund_name).strip()

                # Skip empty values or header rows
                if not code or not name:
                    continue
                if code.lower() in ['fund code', 'short name', 'scheme code', 'fund id']:
                    continue
                if name.lower() in ['scheme full name', 'scheme name', 'fund desc']:
                    continue

                # Skip NOTE/footnote rows (e.g., ['NOTE:', '1', 'None of the schemes...'])
                # These have single-digit codes and disclaimer text
                first_col = row.iloc[0] if len(row) > 0 else None
                if first_col is not None and pd.notna(first_col):
                    first_col_str = str(first_col).strip().lower()
                    if first_col_str.startswith('note'):
                        continue
                # Skip rows where fund_code is just a single digit (note numbers)
                if code.isdigit() and len(code) <= 2:
                    continue

                # Strip parenthetical descriptions and scheme descriptions if configured
                # Handles patterns like:
                # - "Fund Name (description)" -> "Fund Name"
                # - "Fund Name An open-ended scheme description..." -> "Fund Name"
                if index_config and index_config.get('strip_parenthetical', False):
                    # Strip parenthetical content first
                    if '(' in name:
                        name = name.split('(')[0].strip()
                    # Strip scheme descriptions that start with "An open-ended", "A [Type] Scheme", etc.
                    # These patterns typically start a new sentence describing the scheme type
                    # Note: Patterns are designed to be safe and only match actual descriptions, not fund name parts
                    description_patterns = [
                        r'\s+An open-ended.*$',  # "An open-ended..." is always a description
                        r'\s+A\s+[A-Z][^.]*\s+Scheme.*$',  # "A [Type] Scheme" pattern
                        r'\s+A\s+Debt\s+.*$',  # "A Debt..." (with space after Debt to avoid matching "Debt Fund")
                        r'\s+A\s+Equity\s+.*$',  # "A Equity..." (with space after Equity to avoid matching "Equity Fund")
                    ]
                    for pattern in description_patterns:
                        # Use DOTALL flag so . matches newlines (fund names may contain \n)
                        name = re.sub(pattern, '', name, flags=re.IGNORECASE | re.DOTALL).strip()
                        if len(name) < 100:  # Stop if we've cleaned it enough
                            break

                    # Normalize whitespace: replace newlines and multiple spaces with single space
                    # This cleans up names that have newlines from the source data
                    name = re.sub(r'\s+', ' ', name).strip()

                fund_mapping[code] = name
                fund_list_ordered.append((code, name))

            logger.info(f"Parsed {len(fund_mapping)} funds from index sheet '{index_sheet_name}'")
            return fund_mapping, fund_list_ordered

        except Exception as e:
            logger.warning(f"Could not parse index sheet '{index_sheet_name}': {e}")
            return {}, []

    def _get_fund_name_and_code(self, df: pd.DataFrame, sheet_name: str, fund_name_source: Optional[Dict[str, Any]]) -> Tuple[str, str]:
        """
        Get fund name and fund code based on configuration.

        Args:
            df: DataFrame of the sheet
            sheet_name: Name of the sheet (used as default fund_code)
            fund_name_source: Configuration for extracting fund name/code from sheet

        Returns:
            Tuple of (fund_code, fund_name)
        """
        fund_code = sheet_name  # Default
        fund_name = sheet_name  # Default

        # If fund_name_source is configured, extract from sheet row
        if fund_name_source and fund_name_source.get('type') == 'sheet_row':
            try:
                # Extract fund_name
                start_row = fund_name_source.get('row', 0)
                name_col = fund_name_source.get('name_col', 1)
                skip_values = fund_name_source.get('skip_values', [])  # Values to skip (e.g., AMC name)
                max_rows_to_check = fund_name_source.get('max_rows_to_check', 5)  # How many rows to check

                # Try rows starting from start_row until we find a valid fund name
                for row_idx in range(start_row, min(start_row + max_rows_to_check, len(df))):
                    row = df.iloc[row_idx]
                    if name_col >= len(row):
                        continue

                    extracted_name = row.iloc[name_col]
                    if not pd.notna(extracted_name) or not str(extracted_name).strip():
                        continue  # Skip blank rows

                    extracted_name_str = str(extracted_name).strip()

                    # Check if this value should be skipped (e.g., AMC name, "Portfolio as on")
                    # Use "in" for partial matching (e.g., skip "Portfolio as on" matches "Portfolio as on March 15, 2025")
                    should_skip = any(skip_val.lower() in extracted_name_str.lower() for skip_val in skip_values)
                    if should_skip:
                        logger.debug(f"Skipping row {row_idx} - matches skip_values: {extracted_name_str}")
                        continue

                    # Reject header column names and invalid fund name candidates
                    if not _is_valid_fund_name_candidate(extracted_name_str):
                        logger.debug(f"Rejected invalid fund name candidate: {extracted_name_str[:50]}...")
                        continue

                    # Apply strip_before if configured (removes text before and including a delimiter)
                    strip_before = fund_name_source.get('strip_before')
                    if strip_before and strip_before in extracted_name_str:
                        extracted_name_str = extracted_name_str.split(strip_before, 1)[-1].strip()

                    # Apply strip_after if configured (removes text after a delimiter)
                    strip_after = fund_name_source.get('strip_after')
                    if strip_after and strip_after in extracted_name_str:
                        extracted_name_str = extracted_name_str.split(strip_after)[0].strip()

                    # Also handle "PORTFOLIO STATEMENT OF" prefix if present
                    if 'PORTFOLIO STATEMENT OF' in extracted_name_str:
                        extracted_name_str = extracted_name_str.split('PORTFOLIO STATEMENT OF', 1)[-1].strip()

                    # Remove content in parentheses if configured (e.g., "(Formerly known as...)")
                    # Strip everything from first '(' to end (handles nested brackets)
                    if fund_name_source.get('strip_brackets', False):
                        extracted_name_str = extracted_name_str.split('(', 1)[0].strip()

                    # Strip everything after a dash/hyphen if configured (description text)
                    # Only strip if result is >= min_length chars (avoids truncating names like "IIFL- LIQUID FUND")
                    if fund_name_source.get('strip_after_dash', False):
                        min_length = fund_name_source.get('strip_after_dash_min_length', 15)
                        if ' - ' in extracted_name_str:
                            candidate = extracted_name_str.split(' - ', 1)[0].strip()
                            if len(candidate) >= min_length:
                                extracted_name_str = candidate
                        elif '-' in extracted_name_str and not extracted_name_str.startswith('-'):
                            # Handle single dash (but not if it's at the start)
                            parts = extracted_name_str.split('-', 1)
                            if len(parts) > 1 and len(parts[0].strip()) >= min_length:
                                extracted_name_str = parts[0].strip()

                    # Skip if result is empty or looks like a date/portfolio statement
                    if not extracted_name_str or 'PORTFOLIO STATEMENT' in extracted_name_str.upper():
                        continue

                    fund_name = extracted_name_str
                    logger.debug(f"Extracted fund name from row {row_idx}: {fund_name}")
                    break  # Found valid fund name, stop searching

                # Check if fund_code should be explicitly blank (no_fund_code: true)
                if fund_name_source.get('no_fund_code', False):
                    fund_code = None
                    logger.debug(f"Fund code set to None (no_fund_code: true)")
                # Extract fund_code if code_row/code_col specified
                elif fund_name_source.get('code_row') is not None:
                    code_row = fund_name_source.get('code_row')
                    code_col = fund_name_source.get('code_col', 0)
                    if code_row < len(df):
                        code_row_data = df.iloc[code_row]
                        if code_col < len(code_row_data):
                            extracted_code = code_row_data.iloc[code_col]
                            if pd.notna(extracted_code) and str(extracted_code).strip():
                                extracted_code_str = str(extracted_code).strip()
                                # Reject values that look like date headers (e.g., "PORTFOLIO STATEMENT AS ON...")
                                date_header_keywords = ['PORTFOLIO', 'STATEMENT', 'AS ON', 'PURSUANT']
                                is_date_header = any(kw.lower() in extracted_code_str.lower() for kw in date_header_keywords)
                                if not is_date_header:
                                    fund_code = extracted_code_str
                                    logger.debug(f"Extracted fund code from row {code_row}: {fund_code}")
                                else:
                                    logger.debug(f"Rejected date header as fund code: {extracted_code_str[:50]}...")

            except Exception as e:
                logger.warning(f"Failed to extract fund name/code from sheet row: {e}")

            # Apply fund_code_mapping only for generic sheet names (e.g., "Sheet1")
            generic_sheets = fund_name_source.get('generic_sheet_names', [])
            fund_code_mapping = fund_name_source.get('fund_code_mapping', {})
            if sheet_name in generic_sheets and fund_name in fund_code_mapping:
                fund_code = fund_code_mapping[fund_name]
                logger.debug(f"Applied fund_code_mapping: '{fund_name}' -> '{fund_code}'")

        # Fall back to index sheet mapping for fund_name
        if fund_name == sheet_name and sheet_name in self.fund_index:
            fund_name = self.fund_index[sheet_name]

        return fund_code, fund_name

    def process_all_sheets(self) -> List[Dict[str, Any]]:
        """
        Process all sheets in the Excel file.

        Returns:
            List of all extracted records
        """
        try:
            # Try default engine first (xlrd for xls, openpyxl for xlsx)
            excel_file = None
            engine_used = "default"

            try:
                # Suppress stderr for xlrd formula warnings
                with suppress_stderr():
                    excel_file = pd.ExcelFile(self.excel_path)
                engine_used = "xlrd/openpyxl"
            except Exception as e:
                logger.warning(f"Default engine failed: {e}, trying calamine")

            # If default engine failed or returns 0 sheets, try calamine
            if excel_file is None or len(excel_file.sheet_names) == 0:
                try:
                    excel_file = pd.ExcelFile(self.excel_path, engine='calamine')
                    engine_used = "calamine"
                    logger.info(f"Using calamine engine for: {self.excel_path}")
                except ImportError:
                    logger.warning("calamine engine not available - install python-calamine for Strict Open XML support")
                    if excel_file is None:
                        raise
                except Exception as e:
                    logger.warning(f"calamine fallback failed: {e}")
                    if excel_file is None:
                        raise

            logger.info(f"Loaded Excel file with {len(excel_file.sheet_names)} sheets (engine: {engine_used})")

            # Parse index sheet (fallback if fund_name_source not configured)
            self.fund_index, fund_list_ordered = self.parse_index_sheet(excel_file)

            # Build positional mapping: map data sheets to fund codes by position
            # This handles cases where sheet names don't match fund codes (e.g., "Sheet1" instead of "SAMONF")
            data_sheets = [s for s in excel_file.sheet_names if not should_skip_sheet(s, self.skip_sheets)]
            sheet_to_fund = {}  # Maps sheet_name -> (fund_code, fund_name)
            self.sheet_to_fund = sheet_to_fund  # Save for external access (e.g., date extraction)

            if fund_list_ordered:
                # Use positional mapping when Index sheet is present
                for i, sheet_name in enumerate(data_sheets):
                    if sheet_name in self.fund_index:
                        # Sheet name matches a fund code directly
                        sheet_to_fund[sheet_name] = (sheet_name, self.fund_index[sheet_name])
                    elif i < len(fund_list_ordered):
                        # Use positional mapping: data_sheet[i] -> fund[i]
                        fund_code, fund_name = fund_list_ordered[i]
                        sheet_to_fund[sheet_name] = (fund_code, fund_name)
                        logger.info(f"Positional mapping: sheet '{sheet_name}' -> fund_code '{fund_code}', fund_name '{fund_name}'")

            # Check if fund_name_source is configured in metadata
            fund_name_source = self.metadata_loader.fund_name_source

            # Process each sheet (skip configured sheets)
            for sheet_name in excel_file.sheet_names:
                if should_skip_sheet(sheet_name, self.skip_sheets):
                    logger.debug(f"Skipping sheet: {sheet_name}")
                    continue

                # Read sheet without headers (suppress stderr for xlrd warnings)
                with suppress_stderr():
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

                # Skip leading blank rows if configured
                if fund_name_source and fund_name_source.get('skip_leading_blank_rows', False):
                    first_valid = df.first_valid_index()
                    if first_valid and first_valid > 0:
                        df = df.iloc[first_valid:].reset_index(drop=True)
                        logger.debug(f"Skipped {first_valid} leading blank rows")

                # Determine fund_code and fund_name
                if sheet_name in sheet_to_fund:
                    # Use mapping from Index sheet (positional or direct match)
                    fund_code, fund_name = sheet_to_fund[sheet_name]
                else:
                    # No Index sheet mapping available - extract from sheet or use sheet name
                    fund_code, fund_name = self._get_fund_name_and_code(df, sheet_name, fund_name_source)
                    # Store the mapping for date assignment
                    sheet_to_fund[sheet_name] = (fund_code, fund_name)

                # Process the sheet
                processor = SheetProcessor(self.metadata_loader)
                records = processor.process_sheet(df, sheet_name, fund_name, fund_code)

                self.all_records.extend(records)
                self.missing_headers.update(processor.missing_headers)
                self.unrecognized_headers.update(processor.unrecognized_headers)
                self.failed_sheets.update(processor.failed_sheets)

            logger.info(f"Total records extracted: {len(self.all_records)}")
            return self.all_records

        except Exception as e:
            logger.error(f"Error processing Excel file: {e}")
            raise
