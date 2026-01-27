"""
Header Hunter v8.0 - Excel Column Mapping Module

Provides safe, semantic column references that survive schema changes.
Instead of hard-coded column indices, use column names.
"""

from typing import Dict, List, NamedTuple, Optional


class ColumnRef(NamedTuple):
    """Semantic reference to an Excel column."""
    name: str          # Display name (e.g., "Case Size")
    letter: str        # Excel column letter (e.g., "D")
    index: int         # 0-based index (e.g., 3)


class ExcelColumnMap:
    """
    Maps semantic column names to Excel positions.
    
    Usage:
        col_map = ExcelColumnMap(['SKU', 'Product Name', 'Case Size', 'Case Cost'])
        
        # Get column letter by name
        size_col = col_map.get_letter('Case Size')  # Returns 'C' (assuming 0-indexed start)
        
        # Get column index by name
        cost_idx = col_map.get_index('Case Cost')   # Returns 3
    
    Features:
    - Recalculates all positions automatically if columns added/removed
    - Type-safe: raises KeyError with helpful message if column not found
    - Immutable: create once per report, never changes during execution
    """
    
    def __init__(self, column_names: List[str], start_index: int = 0):
        """
        Initialize column map.
        
        Args:
            column_names: Ordered list of column names (e.g., ['SKU', 'Product Name', ...])
            start_index: Starting column index (default 0 = column A)
        """
        self._column_names = column_names.copy()
        self._column_map: Dict[str, ColumnRef] = {}
        
        for idx, col_name in enumerate(column_names, start=start_index):
            col_letter = self._index_to_letter(idx)
            self._column_map[col_name] = ColumnRef(
                name=col_name,
                letter=col_letter,
                index=idx
            )
    
    @staticmethod
    def _index_to_letter(idx: int) -> str:
        """
        Convert 0-based column index to Excel letter (0->A, 1->B, ..., 26->AA).
        
        Args:
            idx: 0-based column index
            
        Returns:
            Excel column letter(s) (e.g., 'A', 'B', 'AA', 'AB')
        """
        result = ""
        idx_copy = idx
        while True:
            result = chr(65 + (idx_copy % 26)) + result
            idx_copy = idx_copy // 26 - 1
            if idx_copy < 0:
                break
        return result
    
    def get_ref(self, column_name: str) -> ColumnRef:
        """
        Get ColumnRef (letter + index) by semantic name.
        
        Args:
            column_name: Column name (must exist in map)
            
        Returns:
            ColumnRef with letter and index
            
        Raises:
            KeyError: If column not found, with helpful message listing available columns
        """
        if column_name not in self._column_map:
            available = ', '.join(self._column_names)
            raise KeyError(
                f"Column '{column_name}' not found. "
                f"Available columns: {available}"
            )
        return self._column_map[column_name]
    
    def get_letter(self, column_name: str) -> str:
        """
        Get Excel column letter by name (e.g., 'Case Size' -> 'D').
        
        Args:
            column_name: Column name
            
        Returns:
            Excel column letter(s)
        """
        return self.get_ref(column_name).letter
    
    def get_index(self, column_name: str) -> int:
        """
        Get 0-based column index by name (e.g., 'Case Size' -> 3).
        
        Args:
            column_name: Column name
            
        Returns:
            0-based column index
        """
        return self.get_ref(column_name).index
    
    def __len__(self) -> int:
        """Return number of columns in map."""
        return len(self._column_names)
    
    def __contains__(self, column_name: str) -> bool:
        """Check if column exists in map."""
        return column_name in self._column_map
    
    def list_columns(self) -> List[str]:
        """Return ordered list of all column names."""
        return self._column_names.copy()
    
    def to_dict(self) -> Dict[str, Dict[str, str]]:
        """Export map as dictionary for debugging."""
        return {
            name: {'letter': ref.letter, 'index': str(ref.index)}
            for name, ref in self._column_map.items()
        }


class LocationColumnGroup:
    """
    Represents a group of columns for one location (Hill, Valley, Jasper).
    
    Example:
        metrics_names = ['Status', 'Stock', 'Buy(Cs)', 'Incoming', ...]
        hill_group = LocationColumnGroup('Hill', metrics_names, start_index=10)
        
        status_col = hill_group.get_letter('Status')  # 'K' (assuming 0-indexed)
        buy_col = hill_group.get_letter('Buy(Cs)')    # 'M'
    """
    
    def __init__(self, location_name: str, metric_names: List[str], start_index: int):
        """
        Initialize location column group.
        
        Args:
            location_name: Display name (e.g., 'Hill', 'Valley')
            metric_names: Column names for this location (e.g., ['Status', 'Stock', ...])
            start_index: Starting column index for this location
        """
        self.location_name = location_name
        self._col_map = ExcelColumnMap(metric_names, start_index)
    
    def get_letter(self, metric_name: str) -> str:
        """
        Get column letter for a metric at this location.
        
        Args:
            metric_name: Name of the metric (e.g., 'Status', 'Stock')
            
        Returns:
            Excel column letter(s)
        """
        return self._col_map.get_letter(metric_name)
    
    def get_index(self, metric_name: str) -> int:
        """
        Get column index for a metric at this location.
        
        Args:
            metric_name: Name of the metric
            
        Returns:
            0-based column index
        """
        return self._col_map.get_index(metric_name)
    
    def __len__(self) -> int:
        """Return number of metrics in this location group."""
        return len(self._col_map)
    
    def __repr__(self) -> str:
        return f"LocationColumnGroup({self.location_name}, {self._col_map.list_columns()})"

