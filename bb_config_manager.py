"""
ConfigManager for Blaze Buddy

Handles creation, reading, validation, and mutation of dynamic location configurations.
This is the linchpin for transforming Header Hunter from hardcoded stores to n-location SaaS.

Philosophy:
- Config is the single source of truth for all locations
- Reads are immutable; mutations are explicit and validated
- Failures are loud and early
- Backward compatibility is maintained via versioning
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict, field
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ============================================================================
# DATACLASSES: The schema for our config
# ============================================================================

@dataclass
class Location:
    """
    Represents a single store location.
    
    Attributes:
        id: Unique identifier (e.g., 'hill', 'valley', 'jasper')
        name: Display name (e.g., 'Hill Store')
        store_code: Internal code used by inventory systems
        data_path: Path to location's inventory files
        is_active: Whether this location is currently operational
        created_at: ISO timestamp of creation
        metadata: Extensible dict for future fields (address, region, etc.)
    """
    id: str
    name: str
    store_code: str
    data_path: str
    is_active: bool = True
    created_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dict for JSON serialization."""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Location':
        """Create from dict (e.g., from JSON)."""
        return cls(**data)
    
    def validate(self) -> None:
        """
        Validate location data. Raises ValueError if invalid.
        """
        if not self.id or not isinstance(self.id, str):
            raise ValueError("Location 'id' must be a non-empty string")
        if not self.name or not isinstance(self.name, str):
            raise ValueError("Location 'name' must be a non-empty string")
        if not self.store_code or not isinstance(self.store_code, str):
            raise ValueError("Location 'store_code' must be a non-empty string")
        if not self.data_path or not isinstance(self.data_path, str):
            raise ValueError("Location 'data_path' must be a non-empty string")
        if not isinstance(self.is_active, bool):
            raise ValueError("Location 'is_active' must be a boolean")


@dataclass
class ConfigSchema:
    """
    Root configuration object.
    
    Attributes:
        version: Config schema version for migration purposes
        locations: List of Location objects
        created_at: ISO timestamp of config creation
        updated_at: ISO timestamp of last update
        metadata: Extensible metadata (company name, environment, etc.)
    """
    version: str = "1.0"
    locations: List[Location] = field(default_factory=list)
    created_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dict for JSON serialization."""
        return {
            'version': self.version,
            'locations': [loc.to_dict() for loc in self.locations],
            'created_at': self.created_at,
            'updated_at': self.updated_at,
            'metadata': self.metadata,
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ConfigSchema':
        """Create from dict (e.g., from JSON)."""
        locations = [Location.from_dict(loc) for loc in data.get('locations', [])]
        return cls(
            version=data.get('version', '1.0'),
            locations=locations,
            created_at=data.get('created_at', datetime.utcnow().isoformat()),
            updated_at=data.get('updated_at', datetime.utcnow().isoformat()),
            metadata=data.get('metadata', {}),
        )
    
    def validate(self) -> None:
        """Validate the entire config. Raises ValueError if invalid."""
        if not self.version:
            raise ValueError("Config 'version' is required")
        
        if not isinstance(self.locations, list):
            raise ValueError("Config 'locations' must be a list")
        
        if len(self.locations) == 0:
            raise ValueError("Config must contain at least one location")
        
        # Validate each location
        for loc in self.locations:
            loc.validate()
        
        # Check for duplicate IDs
        ids = [loc.id for loc in self.locations]
        if len(ids) != len(set(ids)):
            raise ValueError("Duplicate location IDs detected")


# ============================================================================
# CONFIGMANAGER: The main API
# ============================================================================

class ConfigManager:
    """
    Manages all operations on the Blaze Buddy configuration.
    
    This is the single source of truth for location data. All mutations go
    through this class to ensure validation and logging.
    
    Usage:
        config = ConfigManager('blaze_buddy_config.json')
        for location in config.get_locations():
            print(location['name'])
        
        config.add_location(Location(id='newstore', name='New Store', ...))
        config.save()
    """
    
    def __init__(self, config_path: str = 'blaze_buddy_config.json'):
        """
        Initialize the ConfigManager.
        
        Args:
            config_path: Path to JSON config file. Creates default if missing.
        """
        self.config_path = Path(config_path)
        self._config: Optional[ConfigSchema] = None
        
        # Load existing config or bootstrap default
        if self.config_path.exists():
            self.load()
            logger.info(f"Loaded config from {self.config_path}")
        else:
            self._bootstrap_default()
            logger.info(f"Created default config at {self.config_path}")
    
    def _bootstrap_default(self) -> None:
        """
        Create a default configuration with example locations.
        This is the skeleton for first-time users.
        """
        default_locations = [
            Location(
                id='hill',
                name='Hill Store',
                store_code='HILL001',
                data_path='./data/hill',
                is_active=True,
            ),
            Location(
                id='valley',
                name='Valley Store',
                store_code='VALLEY001',
                data_path='./data/valley',
                is_active=True,
            ),
            Location(
                id='jasper',
                name='Jasper Store',
                store_code='JASPER001',
                data_path='./data/jasper',
                is_active=True,
            ),
        ]
        
        self._config = ConfigSchema(
            version='1.0',
            locations=default_locations,
            metadata={'source': 'blaze_buddy_bootstrap'},
        )
        self.save()
    
    # ========================================================================
    # READ OPERATIONS
    # ========================================================================
    
    def load(self) -> None:
        """
        Load config from JSON file.
        Raises FileNotFoundError, json.JSONDecodeError, or ValueError if invalid.
        """
        try:
            with open(self.config_path, 'r') as f:
                data = json.load(f)
            
            self._config = ConfigSchema.from_dict(data)
            self._config.validate()
            logger.info(f"Config loaded and validated: {len(self._config.locations)} locations")
        
        except json.JSONDecodeError as e:
            logger.error(f"JSON decode error in {self.config_path}: {e}")
            raise
        except ValueError as e:
            logger.error(f"Config validation error: {e}")
            raise
    
    def get_locations(self, active_only: bool = True) -> List[Dict[str, Any]]:
        """
        Get all locations as dictionaries (immutable view).
        
        Args:
            active_only: If True, return only active locations
        
        Returns:
            List of location dicts
        """
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        locations = self._config.locations
        if active_only:
            locations = [loc for loc in locations if loc.is_active]
        
        return [loc.to_dict() for loc in locations]
    
    def get_location_by_id(self, location_id: str) -> Optional[Dict[str, Any]]:
        """Get a single location by ID. Returns None if not found."""
        locations = self.get_locations(active_only=False)
        for loc in locations:
            if loc['id'] == location_id:
                return loc
        return None
    
    def get_location_ids(self, active_only: bool = True) -> List[str]:
        """Get all location IDs (convenience method for iteration)."""
        return [loc['id'] for loc in self.get_locations(active_only=active_only)]
    
    def get_metadata(self) -> Dict[str, Any]:
        """Get global metadata."""
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        return dict(self._config.metadata)
    
    # ========================================================================
    # WRITE OPERATIONS
    # ========================================================================
    
    def add_location(self, location: Location) -> None:
        """
        Add a new location to the config.
        
        Args:
            location: Location object to add
        
        Raises:
            ValueError: If location is invalid or ID already exists
        """
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        location.validate()
        
        # Check for duplicate ID
        if any(loc.id == location.id for loc in self._config.locations):
            raise ValueError(f"Location with ID '{location.id}' already exists")
        
        self._config.locations.append(location)
        self._config.updated_at = datetime.utcnow().isoformat()
        logger.info(f"Added location: {location.id} ({location.name})")
    
    def remove_location(self, location_id: str) -> None:
        """
        Remove a location by ID.
        
        Args:
            location_id: ID of location to remove
        
        Raises:
            ValueError: If location not found or is the last location
        """
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        if len(self._config.locations) <= 1:
            raise ValueError("Cannot remove the last location")
        
        original_count = len(self._config.locations)
        self._config.locations = [
            loc for loc in self._config.locations if loc.id != location_id
        ]
        
        if len(self._config.locations) == original_count:
            raise ValueError(f"Location '{location_id}' not found")
        
        self._config.updated_at = datetime.utcnow().isoformat()
        logger.info(f"Removed location: {location_id}")
    
    def update_location(self, location_id: str, **kwargs) -> None:
        """
        Update a location's fields.
        
        Args:
            location_id: ID of location to update
            **kwargs: Fields to update (name, store_code, data_path, is_active, etc.)
        
        Raises:
            ValueError: If location not found or update is invalid
        """
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        location = None
        for loc in self._config.locations:
            if loc.id == location_id:
                location = loc
                break
        
        if location is None:
            raise ValueError(f"Location '{location_id}' not found")
        
        # Update allowed fields
        allowed_fields = {'name', 'store_code', 'data_path', 'is_active', 'metadata'}
        for key, value in kwargs.items():
            if key not in allowed_fields:
                raise ValueError(f"Cannot update field '{key}'")
            setattr(location, key, value)
        
        # Re-validate after update
        try:
            location.validate()
        except ValueError as e:
            logger.error(f"Validation error after update: {e}")
            raise
        
        self._config.updated_at = datetime.utcnow().isoformat()
        logger.info(f"Updated location: {location_id}")
    
    def deactivate_location(self, location_id: str) -> None:
        """
        Soft-delete: Mark a location as inactive (doesn't remove it).
        Useful for audit trails and historical data.
        """
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        active_locations = [loc for loc in self._config.locations if loc.is_active]
        if len(active_locations) <= 1:
            raise ValueError("Cannot deactivate the last active location")
        
        self.update_location(location_id, is_active=False)
        logger.info(f"Deactivated location: {location_id}")
    
    def set_metadata(self, key: str, value: Any) -> None:
        """Set global metadata."""
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        self._config.metadata[key] = value
        self._config.updated_at = datetime.utcnow().isoformat()
        logger.info(f"Updated metadata: {key}")
    
    # ========================================================================
    # PERSISTENCE
    # ========================================================================
    
    def save(self) -> None:
        """
        Write the current config to JSON file.
        Atomically (write to temp, then rename) to avoid corruption.
        """
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        self._config.validate()
        
        # Write to temp file, then atomically rename
        temp_path = self.config_path.with_suffix('.json.tmp')
        
        try:
            with open(temp_path, 'w') as f:
                json.dump(self._config.to_dict(), f, indent=2)
            
            # Atomic rename
            temp_path.replace(self.config_path)
            logger.info(f"Config saved to {self.config_path}")
        
        except Exception as e:
            logger.error(f"Error saving config: {e}")
            # Clean up temp file if it exists
            if temp_path.exists():
                temp_path.unlink()
            raise
    
    def export_json(self) -> str:
        """Export config as JSON string (for API responses, backups)."""
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        return json.dumps(self._config.to_dict(), indent=2)
    
    # ========================================================================
    # UTILITY
    # ========================================================================
    
    def info(self) -> Dict[str, Any]:
        """Get debug info about current config state."""
        if self._config is None:
            raise RuntimeError("Config not loaded. Call load() first.")
        
        return {
            'version': self._config.version,
            'total_locations': len(self._config.locations),
            'active_locations': len([loc for loc in self._config.locations if loc.is_active]),
            'created_at': self._config.created_at,
            'updated_at': self._config.updated_at,
            'config_path': str(self.config_path),
        }


# ============================================================================
# STANDALONE UTILITY FUNCTIONS
# ============================================================================

def validate_config_file(filepath: str) -> tuple[bool, str]:
    """
    Validate a config file WITHOUT loading it into memory.
    
    Returns (is_valid, error_message).
    
    Useful for CLI checks before operations.
    """
    try:
        with open(filepath, 'r') as f:
            data = json.load(f)
        
        schema = ConfigSchema.from_dict(data)
        schema.validate()
        
        return True, "Config is valid"
    
    except FileNotFoundError:
        return False, f"File not found: {filepath}"
    
    except json.JSONDecodeError as e:
        return False, f"JSON decode error: {e}"
    
    except ValueError as e:
        return False, f"Validation error: {e}"


# ============================================================================
# EXAMPLE USAGE & TESTING
# ============================================================================

if __name__ == '__main__':
    """
    Quick demo of ConfigManager capabilities.
    Run with: python hh_config_manager.py
    """
    
    # Initialize (creates default config if missing)
    config = ConfigManager('test_blaze_buddy_config.json')
    
    print("=" * 70)
    print("CONFIG MANAGER DEMO")
    print("=" * 70)
    
    # 1. View all locations
    print("\n1. All Locations:")
    for loc in config.get_locations():
        print(f"   - {loc['id']}: {loc['name']} ({loc['store_code']})")
    
    # 2. Get info
    print(f"\n2. Config Info:")
    for key, value in config.info().items():
        print(f"   {key}: {value}")
    
    # 3. Add a new location
    print("\n3. Adding New Location (Denver)...")
    new_location = Location(
        id='denver',
        name='Denver Store',
        store_code='DENVER001',
        data_path='./data/denver',
        is_active=True,
    )
    config.add_location(new_location)
    config.save()
    print(f"   Locations now: {config.get_location_ids()}")
    
    # 4. Update a location
    print("\n4. Updating Location (Valley -> Valley Mega Store)...")
    config.update_location('valley', name='Valley Mega Store')
    config.save()
    valley = config.get_location_by_id('valley')
    print(f"   Updated name: {valley['name']}")
    
    # 5. Deactivate a location
    print("\n5. Deactivating Denver...")
    config.deactivate_location('denver')
    config.save()
    print(f"   Active locations: {config.get_location_ids(active_only=True)}")
    print(f"   All locations: {config.get_location_ids(active_only=False)}")
    
    # 6. Export for API
    print("\n6. Exported JSON (for API):")
    print(config.export_json()[:200] + "...\n")
    
    # Cleanup
    os.remove('test_blaze_buddy_config.json')
    print("✓ Demo complete. Test file cleaned up.")
