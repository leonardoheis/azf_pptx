class AppError(Exception):
    """Base class for application exceptions."""

    def __init__(self, message, status_code=500):
        super().__init__(message)
        self.status_code = status_code


class ValidationError(AppError):
    """Raised when input data is missing or invalid."""

    def __init__(self, message):
        super().__init__(message, status_code=400)


class TemplateError(AppError):
    """Raised when the PPTX template is missing placeholders or is invalid."""

    def __init__(self, message):
        super().__init__(message, status_code=500)
