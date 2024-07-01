class PDFError(Exception):
    """Custom exception for PDF processing errors to add context."""
    def __init__(self, message, original_exception=None):
        super().__init__(f"{message}: {str(original_exception)}")
        self.original_exception = original_exception
