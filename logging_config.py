"""
Header Hunter v8.0 - Logging Configuration Module

Provides structured logging to both file and console.
Logs are written to Documents/HeaderHunter/Logs/ directory.
"""

import logging
import os
from pathlib import Path
from datetime import datetime
from typing import Optional


def setup_logging(log_level: int = logging.INFO, log_to_file: bool = True) -> logging.Logger:
    """
    Configure application-wide logging.
    
    Creates log files in Documents/HeaderHunter/Logs/ directory.
    Logs are rotated daily (new file each day).
    
    Args:
        log_level: Logging level (logging.DEBUG, INFO, WARNING, ERROR)
        log_to_file: Whether to write logs to file (default True)
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger('HeaderHunter')
    logger.setLevel(log_level)
    
    # Prevent duplicate handlers if called multiple times
    if logger.handlers:
        return logger
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console handler (always enabled)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler (if enabled)
    if log_to_file:
        try:
            # Create log directory in Documents/HeaderHunter/Logs/
            docs_path = Path.home() / 'Documents' / 'HeaderHunter' / 'Logs'
            docs_path.mkdir(parents=True, exist_ok=True)
            
            # Log file name: HeaderHunter_YYYY-MM-DD.log
            log_filename = docs_path / f"HeaderHunter_{datetime.now().strftime('%Y-%m-%d')}.log"
            
            file_handler = logging.FileHandler(log_filename, encoding='utf-8')
            file_handler.setLevel(log_level)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
            logger.info(f"Logging to file: {log_filename}")
        except Exception as e:
            # If file logging fails, continue with console only
            logger.warning(f"Could not set up file logging: {e}")
    
    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Get a logger instance for a specific module.
    
    Args:
        name: Logger name (usually __name__). If None, returns root logger.
        
    Returns:
        Logger instance
    """
    if name:
        return logging.getLogger(f'HeaderHunter.{name}')
    return logging.getLogger('HeaderHunter')


# Initialize default logger on import
default_logger = setup_logging()

