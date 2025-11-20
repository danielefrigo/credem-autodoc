from typer import Option, Typer
from typer.main import get_command
from typing import Annotated
from src.generate_doc import generate_doc
from src.generate_test_plan import generate_test_plan


app = Typer(name="autodoc")

generate = Typer(name="generate")

app.add_typer(generate)


DBT_PATH = Annotated[
    str,
    Option("--dbt-path", help="The path to the dbt project for which the documentation has to be generated."),
]


@generate.command()
def doc(
    dbt_path: DBT_PATH
):
    generate_doc(
        dbt_path=dbt_path
    )

@generate.command()
def testplan(
    dbt_path: DBT_PATH
):
    generate_test_plan(
        dbt_path=dbt_path
    )


typer_click_object = get_command(app)
