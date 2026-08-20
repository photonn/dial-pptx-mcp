"""
Core utility functions for PowerPoint MCP Server.
Basic operations and error handling.
"""
import logging
from typing import Any, Callable, List, Tuple, Optional

from logging_utils import get_logger, flatten

logger = get_logger("utils.core")


def try_multiple_approaches(operation_name: str, approaches: List[Tuple[Callable, str]]) -> Tuple[Any, Optional[str]]:
    """
    Try multiple approaches to perform an operation, returning the first successful result.
    
    Args:
        operation_name: Name of the operation for error reporting
        approaches: List of (approach_func, description) tuples to try
        
    Returns:
        Tuple of (result, None) if any approach succeeded, or (None, error_messages) if all failed
    """
    error_messages = []
    
    for approach_func, description in approaches:
        try:
            result = approach_func()
            return result, None
        except Exception as e:
            logger.debug("approach_failed operation=%s approach=%s error=%s",
                         operation_name, description, flatten(str(e)))
            error_messages.append(f"{description}: {str(e)}")

    logger.warning("operation_failed operation=%s approaches_tried=%d",
                   operation_name, len(approaches))
    return None, f"Failed to {operation_name} after trying multiple approaches: {'; '.join(error_messages)}"


def safe_operation(operation_name: str, operation_func: Callable, error_message: Optional[str] = None, *args, **kwargs) -> Tuple[Any, Optional[str]]:
    """
    Execute an operation safely with standard error handling.
    
    Args:
        operation_name: Name of the operation for error reporting
        operation_func: Function to execute
        error_message: Custom error message (optional)
        *args, **kwargs: Arguments to pass to the operation function
        
    Returns:
        A tuple (result, error) where error is None if operation was successful
    """
    try:
        result = operation_func(*args, **kwargs)
        return result, None
    except ValueError as e:
        error_msg = error_message or f"Invalid input for {operation_name}: {str(e)}"
        logger.warning("operation_failed operation=%s kind=value_error error=%s",
                       operation_name, flatten(str(e)))
        return None, error_msg
    except TypeError as e:
        error_msg = error_message or f"Type error in {operation_name}: {str(e)}"
        logger.warning("operation_failed operation=%s kind=type_error error=%s",
                       operation_name, flatten(str(e)))
        return None, error_msg
    except Exception as e:
        error_msg = error_message or f"Failed to execute {operation_name}: {str(e)}"
        logger.warning("operation_failed operation=%s kind=%s error=%s",
                       operation_name, type(e).__name__, flatten(str(e)),
                       exc_info=logger.isEnabledFor(logging.DEBUG))
        return None, error_msg