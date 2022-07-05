# A module for all of the graphs in the dashboard

from openpyxl.chart import LineChart, StockChart, Reference # Used to make the charts
from openpyxl.chart.axis import ChartLines # Used to display the lines in the stock chart
from openpyxl.chart.updown_bars import UpDownBars # Used to display the bars in the stock chart
from openpyxl.chart.data_source import NumData, NumVal # Used for "dummy" data in the stock chart
from openpyxl.chart.shapes import GraphicalProperties # Used for changing background color

# Displays a graph that shows the closing values of each day
def standard_stock_graph(wb_obj, graph_data, ticker_symbol, CHART_WIDTH, CHART_HEIGHT, chart_color):
    # Get the data worksheet the chart will be based on
    data_worksheet = wb_obj[graph_data[0]]

    # Convert to date
    for i in range(2, graph_data[1] + 1):
        # No need to have all the date names
        if i % 5 == 0:
            current_value = data_worksheet[f'A{i}'].value
            data_worksheet[f'A{i}'].value = f'=TEXT({current_value}, "d mmm yy")'
        else:
            data_worksheet[f'A{i}'].value = "" # Remove all other dates

    # Create the chart
    line_chart = LineChart()
    line_chart.title = f"{ticker_symbol} Recent Closing Prices"
    line_chart.x_axis.title = ""
    line_chart.y_axis.title = ""
    line_chart.width = CHART_WIDTH
    line_chart.height = CHART_HEIGHT

    line_chart.plot_area.graphicalProperties = GraphicalProperties(solidFill=chart_color)

    # Add values and dates
    values = Reference(data_worksheet, min_col=2, min_row=1, max_col=graph_data[2], max_row=graph_data[1])
    labels = Reference(data_worksheet, min_col=1, min_row=2, max_col=1, max_row=graph_data[1])

    # Add all values/ data and the proper titles
    line_chart.add_data(values, titles_from_data=True)
    line_chart.set_categories(labels)

    return line_chart

# Displays a graph that displays the open, high, low, and close values of shares traded
def open_high_low_close_graph(wb_obj, graph_data, CHART_WIDTH, CHART_HEIGHT, chart_color):
    # Get the data worksheet the chart will be based on
    data_worksheet = wb_obj[graph_data[0]]

    for i in range(2, graph_data[1] + 1):
        current_value = data_worksheet[f'A{i}'].value
        data_worksheet[f'A{i}'].value = f'=TEXT({current_value}, "ddd, m/d")'

    # Create the chart
    stock_chart = StockChart()
    stock_chart.title = "Open-High-Low-Close"
    stock_chart.width = CHART_WIDTH
    stock_chart.height = CHART_HEIGHT

    stock_chart.plot_area.graphicalProperties = GraphicalProperties(solidFill=chart_color)

    # Add values and dates
    values = Reference(data_worksheet, min_col=2, min_row=1, max_col=graph_data[2], max_row=graph_data[1])
    labels = Reference(data_worksheet, min_col=1, min_row=2, max_col=1, max_row=graph_data[1])

    # Add all values/ data and the proper titles
    stock_chart.add_data(values, titles_from_data=True)
    stock_chart.set_categories(labels)

    for s in stock_chart.series:
        s.graphicalProperties.line.noFill = True

    stock_chart.hiLowLines = ChartLines()
    stock_chart.upDownBars = UpDownBars()

    # Due to an Excel bug, high-low lines will only be shown with this dummy data
    pts = [NumVal(idx=i) for i in range(len(values) - 1)]
    cache = NumData(pt=pts)
    stock_chart.series[-1].val.numRef.numCache = cache

    return stock_chart