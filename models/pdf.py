import datetime
import dateparser

from utils.custom_exceptions import PDFError


class PDF:

    def __init__(self, pdf_path, pdf_name):
        self.pdf_path = pdf_path
        self.pdf_name = pdf_name
        self._total = None
        self._date = None
        self.vendor = None

    @property
    def date(self):
        return self._date

    @date.setter
    def date(self, extracted_date: str):
        """
        Setter for the date property that parses and sets the date in a standardized format.
        """
        try:
            if isinstance(extracted_date, str):
                parsed_date = dateparser.parse(extracted_date)
                if parsed_date:
                    self._date = parsed_date.strftime('%Y-%m-%d')
                else:
                    raise ValueError(f"Unable to parse the date string: '{extracted_date}'")
            elif isinstance(extracted_date, datetime.date):
                self._date = extracted_date.strftime('%Y-%m-%d')
            else:
                raise TypeError(f"Expected a string or a datetime.date object, received {type(extracted_date).__name__}: {extracted_date}")
        except (TypeError, ValueError) as ex:
            raise PDFError(f"Error setting date for PDF at {self.pdf_path}", ex) from ex

    @property
    def total(self):
        return self._total

    @total.setter
    def total(self, extracted_total):
        try:
            extracted_total = float(extracted_total)
            self._total = extracted_total
        except ValueError as ex:
            raise PDFError(f"Error setting total for PDF at {self.pdf_path}. The input '{extracted_total}' could not be converted to a float.", ex) from ex

