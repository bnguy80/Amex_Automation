import sys
import typer
from typing import Optional
from rich.console import Console
from automation.amex_automation_orchestrator import AmexAutomationOrchestrator

app = typer.Typer()
console = Console()


def show_welcome_message():
    console.print("[dark_orange]Welcome to the AMEX Statement Automation CLI![/dark_orange]", justify="center")
    console.print("""[light_cyan1]This tool processes AMEX Statement process by handling both invoices and transaction details.
This automates the extraction of pdf invoice data and matching of transactional data from AMEX statements.
Use this CLI to simplify the AMEX Statement process and save time.[/light_cyan1]""", justify="left")


@app.command(help="Processes the AMEX Statement by handling both invoices and transaction details.")
def process_amex(
        amex_path: str = typer.Option(
            "K:/B_Amex",
            prompt="Please enter the directory path for the AMEX Statement workbook",
            help="Directory path of the AMEX statement workbook."
        ),
        amex_statement: str = typer.Option(
            "Amex Corp Feb'24 - Addisu Turi (IT).xlsx",
            prompt="Please enter the AMEX Statement workbook name",
            help="Name of the AMEX statement workbook."
        ),
        amex_start_date: str = typer.Option(
            "01/21/2024",
            prompt="Please enter the start date for the statement processing (MM/DD/YYYY)",
            help="Start date for the statement processing."
        ),
        amex_end_date: str = typer.Option(
            "2/21/2024",
            prompt="Please enter the end date for the statement processing (MM/DD/YYYY)",
            help="End date for the statement processing."
        ),
        macro_parameter_1: Optional[str] = typer.Option(
            r"K:\t3nas\APPS\\",
            prompt="Enter the first macro parameter if any",
            help="Optional first macro parameter."
        ),
        macro_parameter_2: Optional[str] = typer.Option(
            "[02] Feb 2024",
            prompt="Enter the second macro parameter if any",
            help="Optional second macro parameter."
        )
):
    """Processes the AMEX statement by handling both invoices and transaction details."""
    controller = AmexAutomationOrchestrator(amex_path, amex_statement, amex_start_date, amex_end_date, macro_parameter_1,
                                            macro_parameter_2)
    controller.process_invoices_worksheet()
    controller.process_transaction_details_2_worksheet()


@app.command(help="Placeholder for a second process. Define functionality here.")
def process_2():
    print("Second process executed.")


@app.command(help="Placeholder for a third process. Define functionality here.")
def process_3():
    print("Third process executed.")


if __name__ == "__main__":
    if len(sys.argv) == 1:  # If no command is provided when main.py is run, defaults to --help 6/16/2024
        show_welcome_message()
        sys.argv.append("--help")
    app()
