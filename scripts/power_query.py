"""Native Excel Power Query helpers. Pass pywin32 Excel .api objects."""
import re

XL_SRC_EXTERNAL = 0
XL_CMD_SQL = 2


def add_query(book, name, formula, description=""):
    return book.Queries.Add(name, formula, description)


def get_query(book, name):
    return book.Queries.Item(name)


def set_formula(book, name, formula):
    query = get_query(book, name)
    query.Formula = formula
    return query


def delete_query(book, name, *, delete_connection=False):
    get_query(book, name).Delete(delete_connection)


def refresh_query(book, name):
    get_query(book, name).Refresh()
    book.Application.CalculateUntilAsyncQueriesDone()


def _mashup_source(name):
    # Restrict binding names so neither the OLE DB nor SQL string needs escaping.
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_ ]*", name):
        raise ValueError("Use a query name containing letters, digits, underscores, or spaces.")
    connection = (
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
        f"Location={name};Extended Properties=\"\""
    )
    return connection, f"SELECT * FROM [{name}]"


def load_worksheet(book, name, sheet, destination):
    """Create a native external ListObject at a caller-verified clear destination."""
    get_query(book, name)
    connection, command = _mashup_source(name)
    table = sheet.ListObjects.Add(
        XL_SRC_EXTERNAL, connection, False, 1, sheet.Range(destination)
    )
    query_table = table.QueryTable
    query_table.CommandType = XL_CMD_SQL
    query_table.CommandText = command
    refresh_worksheet(query_table)
    return table


def refresh_worksheet(query_table):
    if not query_table.Refresh(False):
        raise RuntimeError("Excel cancelled or failed to complete the worksheet refresh.")
    if query_table.FetchedRowOverflow:
        raise RuntimeError("Query results exceeded the Excel worksheet row limit.")


def load_model(book, name, connection_name):
    """Create and refresh a model connection. Verify its table and data separately."""
    get_query(book, name)
    connection, command = _mashup_source(name)
    model_connection = book.Connections.Add2(
        connection_name, "", connection, command, XL_CMD_SQL, True, False
    )
    model_connection.Refresh()
    book.Application.CalculateUntilAsyncQueriesDone()
    return model_connection
