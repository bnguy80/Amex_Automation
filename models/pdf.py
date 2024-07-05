import datetime
import dateparser

from utils.custom_exceptions import PDFError

from dataclasses import dataclass, field


@dataclass
class PDF:
    pdf_path: str
    pdf_name: str
    # init=False means
    # that it shouldn't be settable through the constructor;
    # it needs to be set after the object has been constructed 7/4/2024
    _total: float = field(default=None, init=False)
    _date: str = field(default=None, init=False)
    vendor: str = field(default=None, init=False)

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

    @property
    def date(self):
        return self._date

    @date.setter
    def date(self, extracted_date: str):
        """
        This method sets the date of the PDF to the provided extracted_date.
        The extracted_date can be either a string or a datetime.date object.
        If it is a string, it will be parsed using the dateparser library.
        If the parsing fails, a ValueError will be raised.
        If the extracted_date is already a datetime.date object, it will be used directly.
        If the extracted_date is of any other type, a TypeError will be raised.

        The date will be stored in the self._date attribute of the class instance,
        formatted as "YYYY-MM-DD".

        If any errors occur during date setting,
        a PDFError will be raised with an error message including the path to the PDF file
        where the date setting was attempted.

        :param extracted_date: The date string to set as the date of the PDF.
        :return: None

        Example usage
        ```
        pdf = PDF()
        pdf.date = "2022-01-01"
        Pdf.date = datetime.date(2022, 1, 1)
        ```
        """
        try:
            if isinstance(extracted_date, str):
                parsed_date = dateparser.parse(extracted_date)
                if not parsed_date:
                    raise ValueError(f"Unable to parse the date string: '{extracted_date}'")
            elif isinstance(extracted_date, datetime.date):
                parsed_date = extracted_date
            else:
                raise TypeError(f"Expected a string or a datetime.date object, received {type(extracted_date).__name__}: {extracted_date}")

            self._date = parsed_date.strftime("%Y-%m-%d")
        except (TypeError, ValueError) as ex:
            raise PDFError(f"Error setting date for PDF at {self.pdf_path}", ex) from ex

