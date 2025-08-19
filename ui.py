from rich.console import Console
from rich.table import Table
from rich.text import Text
from rich.padding import Padding
from pandas import DataFrame

def print_user_table_clean(data:DataFrame)->None:
    console = Console()
    table = Table(
        show_header=True,
        header_style="bold white on black",
        border_style="white", 
        padding=(0, 0)
    )

    table.add_column("STT", justify="center")
    for col in data.columns:
        table.add_column(str(col), justify="center")

    # Hàng trắng đầu tiên
    top_blank = [Text("") for _ in range(len(data.columns) + 1)]
    table.add_row(*top_blank)

    for i, row in data.iterrows():
        styled_row = [Text(str(i), style="black on white")]
        for cell in row:
            value = str(cell) if str(cell) != "nan" else " "
            styled_row.append(Text(value, style="black on white"))
        table.add_row(*styled_row)

        if i < len(data) - 1:
            table.add_row(*top_blank)
    padded_table = Padding(table, (0, 2)) 
    console.print(padded_table)


