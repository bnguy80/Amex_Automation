class PDFError(Exception):
    """Custom exception for PDF processing errors to add context."""
    def __init__(self, message: str, original_exception=None):
        super().__init__(f"{message}, Original Exception: {str(original_exception)}")
        self.original_exception = original_exception
