"""Exception hierarchy for Ai_Olap ETL."""


class ETLException(Exception):
    """Base ETL exception."""


class ConnectionFailed(ETLException):
    """Could not establish a SQL or COM connection."""


class MappingError(ETLException):
    """Issue with mapping/baserp_storage.json (missing object, stale, malformed)."""


class ExtractionError(ETLException):
    """Failure during the Extract step."""


class TransformError(ETLException):
    """Failure during the Transform step."""


class LoadError(ETLException):
    """Failure during the Load step."""


class ValidationError(ETLException):
    """Pipeline JSON failed schema validation."""


class AcceptanceFailed(ETLException):
    """An acceptance gate (e.g. Глобино-2 etalon) did not pass."""
