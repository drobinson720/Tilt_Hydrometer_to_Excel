import os
import datetime
from pathlib import Path
import win32com.client as win32
import pythoncom
import pywintypes

class ExcelExportClass:

    def __init__(self):
        self.excel = None

    def _start_excel(self):
        try:
            self.excel = win32.DispatchEx("Excel.Application")
            self.excel.DisplayAlerts = False
            self.excel.Visible = False

        except (pythoncom.com_error, pywintypes.com_error) as e:
            print("ERROR: Excel could not be opened.")
            print(f"Details: {e}")
            sys.exit(1)

    def open_file(self, file_path, sheet_name):
        file_path = str(file_path)

        # Start a background excel instance
        self._start_excel()

        # Create workbook if missing
        if not os.path.exists(file_path):
            self.build_workbook(file_path, sheet_name)

        # Open workbook
        wb = self.excel.Workbooks.Open(file_path, ReadOnly=False)
        return wb

    def open_sheet(self, wb, sheet_name):
        for ws in wb.Worksheets:
            if ws.Name == sheet_name:
                return ws

        ws = wb.Worksheets.Add()
        ws.Name = sheet_name
        self.build_sheet_layout(ws)
        return ws

    def build_workbook(self, file_path, sheet_title):
        wb = self.excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = sheet_title

        self.build_sheet_layout(ws)

        wb.SaveAs(file_path, FileFormat=51)
        wb.Close(SaveChanges=False)

        
    # ---------------------------------------------------------
    # Build sheet layout (headers + summary block)
    # ---------------------------------------------------------
    def build_sheet_layout(self, ws):

        # Headers
        headers = ["Timestamp", "Day", "Temp (°F)", "Gravity", "Tilt Name",
                   "RSSI (dBm)", "Last 8hr", "Last 24hr"]
        #ws.Range("A1").Value = headers
        ws.Range(ws.Cells(1, 1), ws.Cells(1, len(headers))).value = headers
        
        # Summary block
        ws.Range("J1").Value = "8hr Ferm Rate (SG/Day)"
        ws.Range("K1").Formula = (
            '=IFERROR((MAXIFS(D:D,G:G,TRUE)-MINIFS(D:D,G:G,TRUE))/'
            '(MAXIFS(A:A,G:G,TRUE)-MINIFS(A:A,G:G,TRUE)), "tbd")'
        )
        ws.Range("K1").NumberFormat = "0.00000"

        ws.Range("J2").Value = "24hr Ferm Rate (SG/Day)"
        ws.Range("K2").Formula = (
            '=IFERROR((MAXIFS(D:D,H:H,TRUE)-MINIFS(D:D,H:H,TRUE))/'
            '(MAXIFS(A:A,H:H,TRUE)-MINIFS(A:A,H:H,TRUE)), "tbd")'
        )
        ws.Range("K2").NumberFormat = "0.00000"
        
        ws.Range("J3").Value = "Progression"
        ws.Range("K3").Formula = ('=IF(K1<>"tbd",IF(K1>K2,"Speeding Up","Slowing Down"), "tbd")')

        ws.Range("J4").Value = "Current ABV %"
        ws.Range("K4").Formula = '=(MAX(D:D)-MIN(D:D))*1.3125'
        ws.Range("K4").NumberFormat = "0.00%"

        ws.Range("J5").Value = "Apparent Attenuation"
        ws.Range("K5").Formula = '=(MAX(D:D)-MIN(D:D))/(MAX(D:D)-1)'
        ws.Range("K5").NumberFormat = "0.0%"

        ws.Range("J6").Value = "Current Sugar (g/12oz)"
        ws.Range("K6").Formula = '=(MIN(D:D)-1)*1000*2.66*0.355'
        ws.Range("K6").NumberFormat = "0.0"

        ws.Range("J8").Value = "Target/Final SG"
        ws.Range("K8").Value = 1.025

        ws.Range("J9").Value = "Remaining Days"
        ws.Range("K9").Formula = '=IFERROR((MIN(D:D)-K8)/(K2),"tbd")'
        ws.Range("K9").NumberFormat = "0.0"

        ws.Range("J10").Value = "Finish Estimate"
        ws.Range("K10").Formula = '=IFERROR(NOW()+K9, "tbd")'
        ws.Range("K10").NumberFormat = "d-mmm-yy h:mm AM/PM"

        ws.Range("J11").Value = "Final ABV %"
        ws.Range("K11").Formula = '=(MAX(D:D)-K8)*1.3125'
        ws.Range("K11").NumberFormat = "0.00%"

        ws.Range("J12").Value = "Final Attenuation"
        ws.Range("K12").Formula = '=(MAX(D:D)-K8)/(MAX(D:D)-1)'
        ws.Range("K12").NumberFormat = "0.0%"

        ws.Range("J13").Value = "Residual Sugar (g/12oz)"
        ws.Range("K13").Formula = '=(K8-1)*1000*2.66*0.355'
        ws.Range("K13").NumberFormat = "0.0"

        ws.Range("J15").Value = "BT Signal Strength (db)"
        ws.Range("K15").Formula = '=IFERROR(AVERAGEIF(G:G,TRUE,F:F),F2)'
        ws.Range("K15").NumberFormat = "0.0"

        ws.Range("J16").Value = "Signal Condition"
        ws.Range("K16").Formula = '=IF(K14<-85,"Weak","Acceptable")'

        # Resize Columns
        col_width = [21.0, 4.3, 8.75, 6.3, 8.75, 10.0, 7.2, 8.2, 5.0, 21.5, 18.0]
        center = -4108
        for i, w in enumerate(col_width, start=1):        
            ws.Columns(i).ColumnWidth = w
            ws.columns(i).HorizontalAlignment = center

        self.create_dual_axis_chart_at_cell(ws, top_left_cell="M1", x_col='B', sg_col='D', temp_col='C', last_row=2000)


    def create_dual_axis_chart_at_cell(self, ws, top_left_cell="M1",
                                      x_col='B', sg_col='D', temp_col='C', header_row=1,
                                      width=450, height=320, last_row=2000):
        XL_LINE = 4
        XL_PRIMARY = 1
        XL_SECONDARY = 2
        XL_UP = -4162
        NO_MARKER = -4142

        # find last used row
        data_start = 2
        x_range = ws.Range(f"{x_col}{data_start}:{x_col}{last_row}")
        sg_range = ws.Range(f"{sg_col}{data_start}:{sg_col}{last_row}")
        temp_range = ws.Range(f"{temp_col}{data_start}:{temp_col}{last_row}")

        cell = ws.Range(top_left_cell)
        left, top = cell.Left, cell.Top

        charts = ws.ChartObjects()
        chart_obj = charts.Add(left, top, width, height)
        chart = chart_obj.Chart

        # ensure a chart type that supports secondary axis
        chart.ChartType = 74

        # remove existing series
        while chart.SeriesCollection().Count > 0:
            chart.SeriesCollection(1).Delete()

        # Add SG series using Add(Source=...) then set XValues and name
        chart.SeriesCollection().Add(Source=sg_range)
        sc_sg = chart.SeriesCollection(1)
        sc_sg.XValues = x_range
        sc_sg.Name = ws.Range(f"{sg_col}{header_row}").Value
        sc_sg.MarkerStyle = NO_MARKER
        sc_sg.Format.Line.Weight = 1.75
        sc_sg.Format.Line.ForeColor.RGB = 0x00FF0000
        sc_sg.AxisGroup = XL_PRIMARY            # Assign to primary axis

        # Add Temp series
        chart.SeriesCollection().Add(Source=temp_range)
        sc_temp = chart.SeriesCollection(2)
        sc_temp.XValues = x_range
        sc_temp.Name = ws.Range(f"{temp_col}{header_row}").Value
        sc_temp.MarkerStyle = NO_MARKER
        sc_temp.Format.Line.Weight = 1.75
        sc_temp.Format.Line.ForeColor.RGB = 255
        sc_temp.AxisGroup = XL_SECONDARY                # assign to secondary axis

        # Axis titles and chart title
        chart.HasTitle = True
        chart.ChartTitle.Text = "Tilt Hydrometer Measurements"
        chart.HasLegend = False
        chart.Axes(2, XL_PRIMARY).HasTitle = True
        chart.Axes(2, XL_PRIMARY).MinimumScale = float(ws.Range("K8").Value)
        chart.Axes(2, XL_PRIMARY).AxisTitle.Text = "Specific Gravity"
        chart.Axes(2, XL_PRIMARY).TickLabels.NumberFormat = "0.000"
        chart.Axes(2, XL_PRIMARY).AxisTitle.Font.Color = 0x00FF0000
        chart.Axes(2, XL_PRIMARY).AxisTitle.Font.Bold = True
        chart.Axes(2, XL_PRIMARY).AxisTitle.Font.Size = 12
        chart.Axes(2, XL_SECONDARY).HasTitle = True
        chart.Axes(2, XL_SECONDARY).AxisTitle.Text = "Temperature"
        chart.Axes(2, XL_SECONDARY).TickLabels.NumberFormat = "0"
        chart.Axes(2, XL_SECONDARY).AxisTitle.Font.Color = 255
        chart.Axes(2, XL_SECONDARY).AxisTitle.Font.Bold = True
        chart.Axes(2, XL_SECONDARY).AxisTitle.Font.Size = 12
        chart.Axes(1, XL_PRIMARY).HasTitle = True
        chart.Axes(1, XL_PRIMARY).AxisTitle.Text = "Days"
        chart.Axes(1, XL_PRIMARY).TickLabels.NumberFormat = "0"
        chart.Axes(1, XL_PRIMARY).AxisTitle.Font.Bold = True
        chart.Axes(1, XL_PRIMARY).AxisTitle.Font.Size = 12

        return chart
        

    # ---------------------------------------------------------
    # Append data row
    # ---------------------------------------------------------
    def export_to_file(self, ws, name, temp_f, gravity, rssi):

        # Excel serial datetime
        excel_time = (
            datetime.datetime.now() - datetime.datetime(1899, 12, 30)
        ).total_seconds() / 86400

        # Find next row
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
        append_row = last_row + 1

        # Formulas
        eight_hr_test = f'=IF(A{append_row}<>"", IF(NOW()-A{append_row} < (8/24), TRUE, FALSE), "")'
        twentyfour_hr_test = f'=IF(A{append_row}<>"", IF(NOW()-A{append_row} < 1, TRUE, FALSE), "")'

        # Write row
        ws.Cells(append_row, 1).Value = excel_time
        ws.Cells(append_row, 1).NumberFormat = "d-mmm-yy h:mm:ss AM/PM"
        
        ws.Cells(append_row, 2).Formula = f'=A{append_row}-MIN(A:A)'
        ws.Cells(append_row, 2).NumberFormat = "0.0"
        
        ws.Cells(append_row, 3).Value = temp_f
        
        ws.Cells(append_row, 4).Value = gravity
        
        ws.Cells(append_row, 5).Value = name
        
        ws.Cells(append_row, 6).Value = rssi
        
        ws.Cells(append_row, 7).Formula = eight_hr_test
        
        ws.Cells(append_row, 8).Formula = twentyfour_hr_test

        for i in range(1, 8):
            center = -4108
            ws.columns(i).HorizontalAlignment = center
