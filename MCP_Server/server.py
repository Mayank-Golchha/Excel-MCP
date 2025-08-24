# example
# Create a new Excel sheet named "give a path name". Add columns "Product", "UnitsSold", "UnitPrice".
# Fill it with 5 example products and sample numbers for UnitsSold and UnitPrice.
# Then calculate "TotalSales" = UnitsSold * UnitPrice, "Tax" = TotalSales * 0.1, and "NetSales" = TotalSales + Tax.
# Finally, compute the total of NetSales.

import numpy
from mcp.server.fastmcp import FastMCP
import pandas as pd
import math

# Create an MCP server
mcp = FastMCP("Excel")

class ExcelMCP:
    def __init__(self):
        self.path = None
        self.workbook = None
        self.function_map_one_var = {
            "abs" : lambda x : abs(x),
            "boolean_not" : lambda x : not x,
            "sin" : lambda x : math.sin(x),
            "cos" : lambda x : math.cos(x),
            "tan" : lambda x : math.tan(x)
        }

        self.function_map_two_var = {
            "compare": lambda x, y: x > y,
            "equal": lambda x, y: x == y,
            "compare_equal": lambda x, y: x >= y,
            "bitwise_right_shift": lambda x, y : x >> y,
            "bitwise_left_shift": lambda x, y: x << y,
            "power": lambda x, y: x ** y,
            "root": lambda x, y: x ** (1/y),
            "log": lambda x, y: math.log(x,y),

        }

    def read_sheet(self,path,sheet : str = None):
        """reads a sheet"""
        self.path = path
        self.workbook = pd.read_excel(path,sheet_name=sheet)

        if isinstance(self.workbook, dict):
            if sheet is None:
                self.workbook = list(self.workbook.values())[0]
            else:
                self.workbook = self.workbook[sheet]

        return self.workbook

    def modify_column(self,column : str,new_column : numpy.ndarray):
        """Modifies a given column with the new column"""
        self.workbook[column] = new_column

    def sum_column(self,column):
        """Sum a column in the sheet"""
        if self.workbook is None:
            raise ValueError("No sheet loaded. Call load() first.")
        return self.workbook[column].sum()

    def get_row(self,i):
        """Returns ith row of the sheet"""
        return self.workbook.iloc[i]

    def get_column_headers(self) -> list:
        """Returns the head of the sheet"""
        return excel.workbook.columns.to_list()

    def get_column(self,column_name : str) -> list:
        """Returns the column with the given column name"""
        return self.workbook[column_name].to_list()

    def add_columns(self,column1 : str,column2 : str) -> list:
        """Returns the sum of two columns column1 + column2"""
        # return (excel.workbook[column1] + excel.workbook[column2]).to_list()
        return self.workbook[column1].add(excel.workbook[column2]).to_list()

    def subtract_columns(self,column1 : str,column2 : str) -> list:
        """Returns the difference of two columns column1 - column2"""
        # return (excel.workbook[column1] - excel.workbook[column2]).to_list()
        return self.workbook[column1].diff(excel.workbook[column2]).to_list()

    def multiply_columns(self,column1 : str,column2 : str) -> list:
        """Returns the product of two columns column1 * column2"""
        # return (excel.workbook[column1] * excel.workbook[column2]).to_list()
        return self.workbook[column1].mul(excel.workbook[column2]).to_list()

    def divide_columns(self,column1 : str,column2 : str) -> list:
        """Returns column1 divide by column2"""
        return self.workbook[column1].div(excel.workbook[column2]).to_list()

    def add(self,column : str,const : float) -> list:
        """Adds a const to all elements of a column ands return it"""
        return self.workbook[column].add(const).to_list()

    def subtract(self,column : str,const : float) -> list:
        """Subtracts a const to all elements of a column ands return it"""
        return self.workbook[column].diff(const).to_list()

    def multiply(self,column : str,const : float) -> list:
        """Multiplies a const to all elements of a column ands return it"""
        return self.workbook[column].mul(const).to_list()

    def divide(self,column : str,const : float) -> list:
        """Divide a const to all elements of a column ands return it"""
        return self.workbook[column].div(const).to_list()



excel = ExcelMCP()
# @mcp.resource("excel://{path}/{sheet}")


@mcp.tool()
def get_sheet(path: str, sheet: str = None) -> str:
    """Loads the Excel sheet"""
    excel.read_sheet(path,sheet)
    return "Sheet Loaded"

@mcp.tool()
def create_new_sheet(path : str,sheet : str = None):
    """Creates a new sheet in the given path"""
    df = pd.DataFrame()
    df.to_excel(path, sheet_name=sheet, index=False)

@mcp.tool()
def columnSum(column : str) -> int:
    """Sum a column in the sheet"""
    return excel.sum_column(column)

@mcp.tool()
def getColumnHeaders() -> list:
    """Returns the list of column headers of the sheet"""
    return excel.get_column_headers()

@mcp.tool()
def getRow(i : int) -> list:
    """Returns the ith row of the sheet"""
    return excel.get_row(i).to_list()

@mcp.tool()
def columnExists(column : str) -> bool:
    """Checks if a column exists in the header"""
    return column in excel.workbook.columns

@mcp.tool()
def getColumn(column_name : str) -> list:
    """Returns the column with the given column name"""
    return excel.get_column(column_name)

@mcp.tool()
def addColumns(column1: str, column2: str) -> list:
    """Returns the sum of two columns column1 + column2"""
    # return (excel.workbook[column1] + excel.workbook[column2]).to_list()
    return excel.workbook[column1].add(excel.workbook[column2]).to_list()

@mcp.tool()
def subtractColumns(column1: str, column2: str) -> list:
    """Returns the difference of two columns column1 - column2"""
    # return (excel.workbook[column1] - excel.workbook[column2]).to_list()
    return excel.workbook[column1].diff(excel.workbook[column2]).to_list()


@mcp.tool()
def multiplyColumns(column1: str, column2: str) -> list:
    """Returns the product of two columns column1 * column2"""
    # return (excel.workbook[column1] * excel.workbook[column2]).to_list()
    return excel.workbook[column1].mul(excel.workbook[column2]).to_list()


@mcp.tool()
def divideColumns(column1: str, column2: str) -> list:
    """Returns column1 divide by column2"""
    return excel.workbook[column1].div(excel.workbook[column2]).to_list()


@mcp.tool()
def add(column: str, const: float) -> list:
    """Adds a const to all elements of a column ands return it"""
    return excel.workbook[column].add(const).to_list()


@mcp.tool()
def subtract(column: str, const: float) -> list:
    """Subtracts a const to all elements of a column ands return it"""
    return excel.workbook[column].sub(const).to_list()


@mcp.tool()
def multiply(column: str, const: float) -> list:
    """Multiplies a const to all elements of a column ands return it"""
    return excel.workbook[column].mul(const).to_list()


@mcp.tool()
def divide(column: str, const: float) -> list:
    """Divide a const to all elements of a column ands return it"""
    return excel.workbook[column].div(const).to_list()

@mcp.tool()
def function_map_one_var(function_name: str, x: float):
    """
       Execute a one-variable function by name.

       Valid function names:
       - "abs": absolute value
       - "boolean_not": logical NOT
       - "sin": sine (radians)
       - "cos": cosine (radians)
       - "tan": tangent (radians)
       """
    if function_name not in excel.function_map_one_var:
        raise ValueError(f"Function '{function_name}' not found")
    return excel.function_map_one_var[function_name](x)


@mcp.tool()
def function_map_two_var(function_name: str, x: float, y: float) -> float:
    """
        Execute a two-variable function by name.

        Valid function names:
        - "compare": returns True if x > y
        - "equal": returns True if x == y
        - "compare_equal": returns True if x >= y
        - "bitwise_right_shift": x >> y
        - "bitwise_left_shift": x << y
        - "power": x ** y
        - "root": x ** (1/y)
        - "log": log base y of x
        """
    if function_name not in excel.function_map_two_var:
        raise ValueError(f"Function '{function_name}' not found")
    return excel.function_map_two_var[function_name](x, y)


@mcp.tool()
def applyUnaryFunctionOnColumn(column : str,function_name : str):
    """Applies a one variable function on a column and returns it"""
    if function_name not in excel.function_map_one_var:
        raise ValueError(f"Function '{function_name}' not found")
    return excel.workbook[column].apply(excel.function_map_one_var[function_name]).to_list()


@mcp.tool()
def applyTwoVariableFunctionOnColumn(column : str,function_name : str,y :float):
    """Applies a two variable function on a column and returns it"""
    if function_name not in excel.function_map_two_var:
        raise ValueError(f"Function '{function_name}' not found")
    return excel.workbook[column].apply(lambda x : excel.function_map_two_var[function_name](x,y)).to_list()


@mcp.tool()
def filterRowsByCondition(column: str, operator: str, value: float) -> list:
    """
    Filter rows from the sheet where the given column satisfies a condition.

    Supported operators: ">", ">=", "<", "<=", "==", "!="

    Returns the updated column as a list.
    """
    if excel.workbook is None:
        raise ValueError("No sheet loaded. Call read_sheet() first.")

    if operator == ">":
        excel.workbook = excel.workbook[excel.workbook[column] <= value]
    elif operator == ">=":
        excel.workbook = excel.workbook[excel.workbook[column] < value]
    elif operator == "<":
        excel.workbook = excel.workbook[excel.workbook[column] >= value]
    elif operator == "<=":
        excel.workbook = excel.workbook[excel.workbook[column] > value]
    elif operator == "==":
        excel.workbook = excel.workbook[excel.workbook[column] != value]
    elif operator == "!=":
        excel.workbook = excel.workbook[excel.workbook[column] == value]
    else:
        raise ValueError(f"Unsupported operator '{operator}'")

    return excel.workbook[column].to_list()


@mcp.tool()
def readRows(n : int):
    """Reads first n rows of the sheet"""
    return excel.workbook.head(n)

@mcp.tool()
def modifyColumn(column: str, new_column: list):
    """
    Modifies a given column with the new column.
    Automatically extends the DataFrame if the new column is longer.
    Does not write to the Excel file; only modifies in RAM.
    """
    # Convert to numpy array internally
    new_column = numpy.array(new_column)

    # Check if we need to extend the DataFrame
    extra_len = len(new_column) - len(excel.workbook)
    if extra_len > 0:
        # Extend existing rows with NaN
        extra_rows = pd.DataFrame({col: [numpy.nan] * extra_len for col in excel.workbook.columns})
        excel.workbook = pd.concat([excel.workbook, extra_rows], ignore_index=True)

    # Assign the new column
    excel.workbook[column] = new_column.tolist()  # Convert back to list for MCP compatibility


@mcp.tool()
def writeSheet():
    """Writes data in the Excel sheet"""
    excel.workbook.to_excel(excel.path)

@mcp.tool()
def range_array(start : int, end : int, steps : int = 1) -> list:
    """Returns a range array given the start end and step"""
    arr = numpy.arange(start, end, steps).tolist()
    return arr

@mcp.tool()
def linspace_array(start: float, end: float, num: int = 50) -> list:
    """Returns an array with 'num' evenly spaced points between start and end."""
    return numpy.linspace(start, end, num).tolist()

@mcp.tool()
def cumulative_array(column : str) -> list:
    """Returns the cumulative sum array of original array"""
    data = excel.get_column(column)
    return numpy.cumsum(data).tolist()

@mcp.tool()
def random_int_array(start : int,end : int,size : int) -> list:
    """Returns a random integer array with given range start to end to choose a number and the size of array"""
    return numpy.random.randint(start, end, size = size).tolist()

@mcp.tool()
def random_array(size : int) -> list:
    """Returns an array of given size with random no between 0 and 1"""
    return numpy.random.rand(size).tolist()

@mcp.tool()
def random_array_with_choice(choices : list,size : int) -> list:
    """Returns an array of given size with random choice of an element from the given choices"""
    return numpy.random.choice(choices,size = size).tolist()


@mcp.tool()
def drop_column(column : str):
    """Delete a column from the sheet"""
    if column in excel.workbook.columns:
        del excel.workbook[column]


@mcp.tool()
def empty_rows() -> dict:
    """Return all rows which have empty entry"""
    df = excel.workbook[excel.workbook.isna().any(axis=1)]
    return df.to_dict(orient="records")

@mcp.tool()
def duplicated_rows() -> dict:
    """Return all duplicated rows"""
    df = excel.workbook.loc[excel.workbook.duplicated()]
    return df.to_dict(orient="records")

@mcp.tool()
def remove_ith_row(i : int):
    """Remove the ith row from the sheet"""
    excel.workbook.drop(excel.workbook.index[i])

@mcp.tool()
def drop_duplicated_rows(inplace: bool = False) -> dict:
    """Remove duplicated rows"""
    df = excel.workbook.drop_duplicates(inplace=inplace)
    return df.to_dict(orient="records")

@mcp.tool()
def rows_with_value(column: str, value) -> dict:
    """Return rows where a specific column matches the given value"""
    df = excel.workbook.loc[excel.workbook[column] == value]
    return df.to_dict(orient="records")


if __name__ == "__main__":
    # uv run mcp dev excel_mcp_server.py
    print("Starting Server...")
    mcp.run("sse")
