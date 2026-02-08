class DataLoadError(Exception):
    """Raised when an uploaded file cannot be parsed."""


class DataValidationError(Exception):
    """Raised when required columns are missing or invalid."""
