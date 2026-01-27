"""
Header Hunter v8.0 - Business Rules Module

Pure functions for inventory status determination.
No side effects (no file I/O, no Excel generation, no GUI callbacks).
Safe for unit testing and dependency injection.
"""

from dataclasses import dataclass
from typing import Dict, NamedTuple, Optional
from datetime import datetime


class StatusRules(NamedTuple):
    """Immutable container for status determination thresholds."""
    hot_velocity: float          # Units/week: minimum for "Hot" status
    reorder_point: float         # Weeks: minimum stock before reorder
    target_wos: float            # Weeks: desired stock level
    dead_wos: float              # Weeks: threshold for "Dead" classification
    dead_on_hand: int            # Units: minimum on-hand for "Dead"
    good_velocity_multiplier: float = 0.25  # Multiplier for "Good" velocity threshold


@dataclass
class InventoryMetrics:
    """Immutable container for a single SKU's metrics."""
    stock: int                   # Current on-hand quantity
    incoming: int                # Pending purchase order quantity
    is_accessory: bool           # Product type flag
    total_units_sold: float      # Total units sold in report period
    report_days: float           # Full duration of report in days
    report_start_date: datetime  # Start date of analysis period
    last_sale_date: Optional[datetime] = None  # Date of most recent sale
    
    def __post_init__(self):
        """Validate metrics are non-negative."""
        if self.stock < 0:
            raise ValueError(f"Stock cannot be negative: {self.stock}")
        if self.incoming < 0:
            raise ValueError(f"Incoming cannot be negative: {self.incoming}")
        if self.total_units_sold < 0:
            raise ValueError(f"Total units sold cannot be negative: {self.total_units_sold}")
        if self.report_days <= 0:
            raise ValueError(f"Report days must be positive: {self.report_days}")

    def calculate_velocity(self) -> float:
        """
        Calculate units per week, adjusting for out-of-stock periods.
        
        Rules:
        1. If stock is 0 and last_sale_date exists, period ends at last_sale_date.
        2. Otherwise, period is the full report_days.
        
        Returns:
            float: Units sold per week (7-day average)
        """
        period_days = self.report_days
        
        if self.stock == 0 and self.last_sale_date:
            # Item is OOS, check if it sold out early
            days_until_last_sale = (self.last_sale_date - self.report_start_date).days
            # Min 1 day to prevent division by zero or negative days
            period_days = max(1.0, float(days_until_last_sale))
            # Cap at report_days (shouldn't exceed, but safe guard)
            period_days = min(period_days, self.report_days)
            
        weeks = period_days / 7.0
        return self.total_units_sold / weeks if weeks > 0 else 0.0


class StatusDeterminer:
    """
    Pure logic for determining inventory status.
    
    Implements the 5-tier status hierarchy:
    1. Zero velocity (New or Cold)
    2. High velocity (Hot or Reorder)
    3. Medium velocity (Good or Reorder)
    4. Low velocity (Dead)
    5. Default (Minimal)
    """
    
    # Status emoji constants (centralized)
    STATUS_NEW = "✨ New"
    STATUS_COLD = "❄️ Cold"
    STATUS_HOT = "🔥 Hot"
    STATUS_REORDER = "🚨 Reorder"
    STATUS_GOOD = "✅ Good"
    STATUS_DEAD = "💀 Dead"
    STATUS_MINIMAL = "➖"
    
    @staticmethod
    def calculate_effective_wos(
        stock: int,
        incoming: int,
        velocity: float,
        silence_threshold: float = 999.0
    ) -> float:
        """
        Calculate effective weeks-on-stock including incoming inventory.
        
        If velocity is zero, returns silence_threshold (prevents division by zero).
        
        Args:
            stock: Current on-hand quantity
            incoming: Pending quantity
            velocity: Units/week sales rate
            silence_threshold: Default WOS when velocity = 0
            
        Returns:
            Weeks of stock available (current + incoming / weekly velocity)
            
        Raises:
            ValueError: If velocity is negative
        """
        if velocity < 0:
            raise ValueError(f"Velocity cannot be negative: {velocity}")
        
        if velocity == 0:
            return silence_threshold
        
        total_available = max(0, stock) + incoming
        return total_available / velocity
    
    @staticmethod
    def determine_status(
        metrics: InventoryMetrics,
        rules: StatusRules
    ) -> str:
        """
        Determine status for a single SKU based on business rules.
        
        Pure function: same inputs always produce same output, no side effects.
        
        Args:
            metrics: InventoryMetrics object with units, dates, stock, etc.
            rules: StatusRules object with business thresholds
            
        Returns:
            Status string with emoji (e.g., "🔥 Hot")
        """
        velocity = metrics.calculate_velocity()
        effective_oh = max(0, metrics.stock)
        incoming = metrics.incoming
        
        # Calculate current WOS based on NEW velocity
        wos = StatusDeterminer.calculate_effective_wos(
            effective_oh, 0, velocity
        )
        
        # Calculate effective WOS including incoming stock
        effective_wos = StatusDeterminer.calculate_effective_wos(
            effective_oh, incoming, velocity
        )
        
        # === TIER 1: ZERO VELOCITY (New or Cold) ===
        if velocity == 0:
            if incoming > 0:
                return StatusDeterminer.STATUS_NEW  # Incoming stock, no demand yet
            elif effective_oh > 0:
                return StatusDeterminer.STATUS_COLD  # Stocked but no sales
            else:
                return StatusDeterminer.STATUS_MINIMAL  # No stock, no demand
        
        # === TIER 2: HIGH VELOCITY ===
        if velocity >= rules.hot_velocity:
            if wos < rules.reorder_point:
                # Current stock is low
                if effective_wos >= rules.reorder_point:
                    return StatusDeterminer.STATUS_GOOD  # Incoming covers need
                else:
                    return StatusDeterminer.STATUS_REORDER  # Critical: must reorder
            else:
                return StatusDeterminer.STATUS_HOT  # Strong sales, adequate stock
        
        # === TIER 3: MEDIUM VELOCITY ===
        # Threshold is multiplier (default 25%) of hot velocity
        good_vel_threshold = rules.hot_velocity * rules.good_velocity_multiplier
        if velocity >= good_vel_threshold:
            if wos < rules.reorder_point:
                if effective_wos >= rules.reorder_point:
                    return StatusDeterminer.STATUS_GOOD  # Incoming covers need
                else:
                    return StatusDeterminer.STATUS_REORDER
            else:
                return StatusDeterminer.STATUS_GOOD  # Steady sales, adequate stock
        
        # === TIER 4: LOW VELOCITY (Dead Stock) ===
        if wos > rules.dead_wos and effective_oh > rules.dead_on_hand:
            return StatusDeterminer.STATUS_DEAD  # High stock, minimal sales
        
        # === TIER 5: DEFAULT ===
        return StatusDeterminer.STATUS_MINIMAL


def clean_currency(val) -> float:
    """
    Convert currency values to float, handling various formats.
    
    Supports parenthetical negatives (1,234.56) -> -1234.56 and standard formats.
    Returns 0.0 if conversion fails.
    
    Args:
        val: Value to clean (any type)
        
    Returns:
        float: Cleaned numeric value, 0.0 if conversion fails
    """
    import re
    import pandas as pd
    
    if pd.isna(val):
        return 0.0
    
    val_str = str(val).strip()
    
    # Handle parenthetical negatives: (1234.56) -> -1234.56
    if val_str.startswith('(') and val_str.endswith(')'):
        val_str = '-' + val_str[1:-1]
    
    # Remove all non-numeric characters except decimal and negative sign
    clean = re.sub(r'[^\d.-]', '', val_str)
    
    try:
        return float(clean) if clean else 0.0
    except ValueError:
        return 0.0


def rules_dict_to_status_rules(rules_dict: Dict, is_accessory: bool = False) -> StatusRules:
    """
    Convert dictionary rules (from config) to StatusRules NamedTuple.
    
    Args:
        rules_dict: Dictionary with keys: hot_velocity, reorder_point, target_wos, dead_wos, dead_on_hand
        is_accessory: Whether to use accessory logic (affects good_velocity_multiplier)
        
    Returns:
        StatusRules NamedTuple
    """
    return StatusRules(
        hot_velocity=rules_dict.get('hot_velocity', 2.0),
        reorder_point=rules_dict.get('reorder_point', 2.5),
        target_wos=rules_dict.get('target_wos', 4.0),
        dead_wos=rules_dict.get('dead_wos', 26),
        dead_on_hand=rules_dict.get('dead_on_hand', 5),
        good_velocity_multiplier=0.25  # Fixed at 25% for now
    )


def metrics_dict_to_inventory_metrics(metrics_dict: Dict) -> InventoryMetrics:
    """
    Convert dictionary metrics (from pandas row) to InventoryMetrics dataclass.
    
    Args:
        metrics_dict: Dictionary with keys: Stock, Incoming_Num, Is_Accessory, 
                     Total_Sold, Report_Days, Start_Date, Last_Sale_Date
        
    Returns:
        InventoryMetrics dataclass
    """
    return InventoryMetrics(
        stock=int(metrics_dict.get('Stock', 0)),
        incoming=int(metrics_dict.get('Incoming_Num', 0)),
        is_accessory=bool(metrics_dict.get('Is_Accessory', False)),
        total_units_sold=float(metrics_dict.get('Total_Sold', 0.0)),
        report_days=float(metrics_dict.get('Report_Days', 30.0)),
        report_start_date=metrics_dict.get('Start_Date', datetime.now()),
        last_sale_date=metrics_dict.get('Last_Sale_Date')
    )


def calculate_soq(metrics: InventoryMetrics, rules: StatusRules, case_size: int) -> int:
    """
    Calculate Suggested Order Quantity (SOQ) in units.
    
    Formula: CEILING(MAX((Velocity * Target_WOS - Stock - Incoming), 0) / Case_Size) * Case_Size
    Returns the quantity in units (multiple of case_size).
    
    Args:
        metrics: InventoryMetrics containing current stock and sales data
        rules: StatusRules containing the target_wos
        case_size: Number of units per case
        
    Returns:
        int: Suggested order quantity in units
    """
    velocity = metrics.calculate_velocity()
    target_stock = velocity * rules.target_wos
    net_need = target_stock - (metrics.stock + metrics.incoming)
    
    if net_need <= 0:
        return 0
        
    # Standard SOQ calculation: round up to nearest case
    import math
    cases_needed = math.ceil(net_need / max(1, case_size))
    return int(cases_needed * case_size)

