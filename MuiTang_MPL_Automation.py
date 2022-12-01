import PyQt5.QtWidgets as qtwidget
import PyQt5.QtGui as qtgui
import PyQt5.QtCore as qtcore
from pandas import pivot, pivot_table
import win32com.client as win32
import os
import time
# from plyer import notification # copy and paste this in hidden import on auto-py-exe --> plyer.platforms.win.notification

app = qtwidget.QApplication([])

class MainWindow(qtwidget.QWidget):
    def __init__(self):
        super().__init__()
        
        # Set window title
        self.setWindowTitle('Mui Tang - MPL Brands')
        
        height = 100
        width = 500
        
        # Set fixed window size
        # self.height(height)
        # self.width(width)
        
        self.result = "None"
        self.display = qtwidget.QLabel(f"Result: {self.result}")
        self.display.setStyleSheet("background-color: #e3e1da;\
                                    border: 1px solid black;\
                                    padding-left: 5px")
        
        self.btn1 = qtwidget.QPushButton("Run 'AP Status'", self)
        self.btn2 = qtwidget.QPushButton("Run 'Accrual JE'", self)
        self.btn1.clicked.connect(self.pivot_table)
        self.btn2.clicked.connect(self.pivot_table)
        
        # Set progam main layout 
        main_layout = qtwidget.QVBoxLayout()
        
        # Create horizontal box for buttons
        sub_layout = qtwidget.QHBoxLayout()
        
        # Add buttons to horizontal box
        sub_layout.addWidget(self.btn1)
        sub_layout.addWidget(self.btn2)
        
        # Add horizontal layout to vertical box layout
        main_layout.addLayout(sub_layout)
        main_layout.addWidget(self.display)
        
        
        self.setLayout(main_layout)
        self.show()
    
    def print_success(self):
        self.display.setStyleSheet("background-color: #73ba59;\
                                        border: 1px solid black;\
                                        padding-left: 5px")
        self.display.setText("Result: Success!")
        
    def print_error(self, e):
        self.display.setStyleSheet("background-color: #f0553a;\
                                        border: 1px solid black;\
                                        padding-left: 5px")
        self.display.setText(f"Result: {e}")
        
    def print_status(self):
        self.display.setStyleSheet("background-color: #e3e1da;\
                                        border: 1px solid black;\
                                        padding-left: 5px")
        self.display.setText(f"Result: Running")
     
     
     # Function for AP Status starts here ===>  
    def pivot_table(self):
        qtcore.QTimer.singleShot(1, lambda: self.print_status())
        
        try:
            # Get time now
            now = time.strftime("%m%d%y_%H%M%S")

            # ntf = notification
            success_status = "SUCCESS!"
            fail_status = "Something went wrong!"

            entries = os.listdir("C:/")
            folderName = "AP_Status"

            # Check if AP_Status folder exists and create it & subfolders if not
            if folderName not in entries:
                os.makedirs(f"C:/{folderName}")
                os.makedirs(f"C:/{folderName}/Import")
                os.makedirs(f"C:/{folderName}/Export")
                open(f"c:/{folderName}/log.txt", "w")

            appFolderList = os.listdir(f"C:/{folderName}")
            appFolderPath = f"C:/{folderName}"
            importFolderList = os.listdir(f"C:/{folderName}/Import")
            importFolderPath = f"C:/{folderName}/Import"
            exportFolderList = os.listdir(f"C:/{folderName}/Export")
            exportFolderPath = f"C:/{folderName}/Export"

            xlFileName = "" 

            for i in appFolderList:
                if ".xlsx" in i:            
                    xlFileName = i
                    # shutil.move(appFolderPath+"/"+ xlFileName, importFolderPath+"/" + now + "_" + xlFileName)
                    
            # launch excel application
            xlApp = win32.Dispatch('Excel.Application')
            xlApp.Visible = True


            xlSheet1 = "AP Status"
            xlSheet2 = "$_PivotTable"
            xlSheet3 = "Count_PivotTable"
            xlSheet4 = "2B_Onboarded"
            xlSheet5 = "Filtered"
            xlSheet6 = "$_Category_MPL"
            xlSheet7 = "$_Category_Patco"

            # Open workbook
            wb = xlApp.Workbooks.Open(f"{appFolderPath}\{xlFileName}")

            # Get original worksheets
            # ws_1 = wb.Worksheets(xlSheet1)
            
            ws_1 = wb.ActiveSheet
            # Create worksheet named 'AP Status'
            ws_1.Name = xlSheet1


            # Create worksheet name '$_PivotTable'
            ws_2 = wb.Sheets.Add()
            ws_2.Name = xlSheet2
            ws_3 = wb.Sheets.Add()
            ws_3.Name = xlSheet3
            ws_4 = wb.Sheets.Add()
            ws_4.Name = xlSheet4
            ws_5 = wb.Sheets.Add()
            ws_5.Name = xlSheet5
            
            # Add sheet '$_Category_MPL'
            ws_6 = wb.Sheets.Add()
            ws_6.Name = xlSheet6
            
            # Add sheet '$_Category_Patco'
            ws_7 = wb.Sheets.Add()
            ws_7.Name = xlSheet7

            # Select the entire columns from A to AP
            selectRange1 = ws_1.Range("A:AP")
            
            # Filter values in column 10th 'I' - column named 'Invoice Status'
            selectRange1.AutoFilter(10, ("Pending AP action", "Pending approval", "Pending payment", "Pending review", "Review matching"), 7)

            # Copy filtered value from Sheet 1 to Sheet 5
            ws_1.Range("A1").CurrentRegion.Copy()
            (wb.Worksheets(xlSheet5).Range("A1")).PasteSpecial(-4104, -4142, False, False)

            # Replace values in column Q "Payment Method"
            ws_5.Range("Q:Q").Replace("Payable", "Onboarded", 1, 2, False, False )
            ws_5.Range("Q:Q").Replace("Unpayable", "Not Onboarded", 1, 2, False, False )


            # reference worksheet
            # wsInvoices = wb.Worksheets("Invoices")
            # wsPivotTable = wb.Worksheets("PivotTable")

            # create pt cache connection
            pt_cache = wb.PivotCaches().Create(1, ws_5.Range("A1").CurrentRegion)

            # Insert pivot tables to sheets
            pt = pt_cache.CreatePivotTable(ws_2.Range("A1"), "$_PivotTable")
            
            pt2 = pt_cache.CreatePivotTable(ws_3.Range("A1"), "Count Pivot Table")
            
            pt3 = pt_cache.CreatePivotTable(ws_6.Range("A1"), "MPL Brands - Onboarded")
            
            pt5 = pt_cache.CreatePivotTable(ws_7.Range("A1"), "Patco Brands - Onboarded")

            # toggle grand totals on/off for rows
            pt.ColumnGrand = True
            pt.RowGrand = True

            pt2.ColumnGrand = True
            pt2.RowGrand = True
            
            pt3.ColumnGrand = True
            pt3.RowGrand = True    
            
            pt5.ColumnGrand = True
            pt5.RowGrand = True

            # change report layout
            pt.RowAxisLayout(0)
            pt2.RowAxisLayout(0)

            def AP_PivotTable(pt):
                field_rows = {}
                field_rows['entity'] = pt.PivotFields("Entity Name")
                field_rows['status'] = pt.PivotFields("Invoice Status")
                field_rows['payment'] = pt.PivotFields("Payment Method")

                # insert row fields
                field_rows['entity'].Orientation = 1
                field_rows['entity'].Position = 1
                field_rows['entity'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['status'].Orientation = 1
                field_rows['status'].Position = 2
                field_rows['status'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['payment'].Orientation = 1
                field_rows['payment'].Position = 3
                field_rows['payment'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                        
                # Adding Invoice Due Date to Pivot Field Column
                field_columns = {}
                field_columns['due date'] = pt.PivotFields("Invoice Due Date")
                field_columns['due date'].Orientation = 2
                field_columns['due date'].Position = 1


                # Grouping column fields for Month, Quarter, Year
                field_columns['due date'].AutoGroup()
                field_columns['due date'].Orientation = 2
                field_columns['due date'].Position = 2
                field_columns['due date'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_columns['due year'] = pt.PivotFields("Years")
                field_columns['due year'].Orientation = 2
                field_columns['due year'].Position = 1
                field_columns['due year'].RepeatLabels = 2
                field_columns['due year'].ShowDetail = 1
                field_columns['due year'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Setting Pivot Field "Quarter" to be hidden
                field_columns['due quarter'] = pt.PivotFields("Quarters")
                field_columns['due quarter'].Orientation = 0
                        
                field_values = {}
                field_values['amount'] = pt.PivotFields("Invoice Amount")


                # insert value fields
                field_values['amount'].Orientation = 4
                field_values['amount'].Function = -4157

                selectRange = ws_2.Range("A:EV")
                selectRange.Font.Name = "Century Gothic"
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
            def Count_PivotTable(pt2):
                field_rows = {}
                field_rows['entity'] = pt2.PivotFields("Entity Name")
                field_rows['status'] = pt2.PivotFields("Invoice Status")
                field_rows['payment'] = pt2.PivotFields("Payment Method")

                # insert row fields
                field_rows['entity'].Orientation = 1
                field_rows['entity'].Position = 1
                field_rows['entity'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['status'].Orientation = 1
                field_rows['status'].Position = 2
                field_rows['status'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['payment'].Orientation = 1
                field_rows['payment'].Position = 3
                field_rows['payment'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Adding Invoice Due Date to Pivot Field Column
                field_columns = {}
                field_columns['due date'] = pt2.PivotFields("Invoice Due Date")
                field_columns['due date'].Orientation = 2
                field_columns['due date'].Position = 1
        

                # Grouping column fields for Month, Quarter, Year
                field_columns['due date'].AutoGroup()
                field_columns['due date'].Orientation = 2
                field_columns['due date'].Position = 2
                field_columns['due date'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_columns['due year'] = pt2.PivotFields("Years")
                field_columns['due year'].Orientation = 2
                field_columns['due year'].Position = 1
                field_columns['due year'].RepeatLabels = 2
                field_columns['due year'].ShowDetail = 1
                field_columns['due year'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Setting Pivot Field "Quarter" to be hidden
                field_columns['due quarter'] = pt2.PivotFields("Quarters")
                field_columns['due quarter'].Orientation = 0
                        
                field_values = {}
                field_values['amount'] = pt2.PivotFields("Invoice Amount")


                # insert value fields
                field_values['amount'].Orientation = 4
                field_values['amount'].Function = -4112

                selectRange = ws_3.Range("A:EV")
                selectRange.Font.Name = "Century Gothic"
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "General"

            def Not_Onboarded(ws_5,ws_4):
                selectRange2 = ws_5.Range("A:AP")
                selectRange2.AutoFilter(17, "Not Onboarded", 7)
                
                # Select column C - column named 'Company Name'
                ws_5.Range("C:C").Copy()
                (ws_4.Range("A1")).PasteSpecial(-4104, -4142, True, False)
                selectRange3 = ws_4.Range("A:A")
                selectRange3.RemoveDuplicates(1,1)
                
                selectRange3.Font.Name = "Century Gothic"
                ws_4.Range("A1").Value = "To Be Onboarded"
                ws_4.Range("A1").Font.Bold = True
                
                selectRange = ws_4.Range("A:A")
                selectRange.Font.Name = "Century Gothic"
                ws_4.Range("A1").Value = "To Be Onboarded"
                
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "General"

            def Category_MPL1(pt3):
                field_rows = {}
                field_rows['invoice_status'] = pt3.PivotFields("Invoice Status")
                field_rows['category'] = pt3.PivotFields("Category")
                field_rows['company_name'] = pt3.PivotFields("Company Name")

                # insert row fields
                field_rows['invoice_status'].Orientation = 1
                field_rows['invoice_status'].Position = 1
                field_rows['invoice_status'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['category'].Orientation = 1
                field_rows['category'].Position = 2
                field_rows['category'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['company_name'].Orientation = 1
                field_rows['company_name'].Position = 3
                field_rows['company_name'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                        
                # Adding Invoice Due Date to Pivot Field Column
                field_columns = {}
                field_columns['due_date'] = pt3.PivotFields("Invoice Due Date")
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 1


                # Grouping column fields for Month, Quarter, Year
                field_columns['due_date'].AutoGroup()
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 2
                field_columns['due_date'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_columns['due_year'] = pt3.PivotFields("Years")
                field_columns['due_year'].Orientation = 2
                field_columns['due_year'].Position = 1
                field_columns['due_year'].RepeatLabels = 2
                field_columns['due_year'].ShowDetail = 1
                field_columns['due_year'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Setting Pivot Field "Quarter" to be hidden
                field_columns['due_quarter'] = pt3.PivotFields("Quarters")
                field_columns['due_quarter'].Orientation = 0
                        
                field_values = {}
                field_values['amount'] = pt3.PivotFields("Invoice Amount")


                # insert value fields
                field_values['amount'].Orientation = 4
                field_values['amount'].Function = -4157
                
                # Insert filter fields
                field_filters= {}
                
                field_filters['entity_name'] = pt3.PivotFields('Entity Name')
                field_filters['entity_name'].Orientation = 3
                
                field_filters['payment_method'] = pt3.PivotFields('Payment Method')
                field_filters['payment_method'].Orientation = 3
                
                # This puts "Entity Name" field on top of "Payment Method"
                field_filters['entity_name'].Position = 2
                
                field_filters['entity_name'].CurrentPage = "MPL Brands Inc (Default)"
                field_filters['payment_method'].CurrentPage = "Onboarded"
                
                selectRange = ws_6.Range("A:EV")
                selectRange.Font.Name = "Century Gothic"
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
            def Category_MPL2(pt4):
                field_rows = {}
                field_rows['invoice_status'] = pt4.PivotFields("Invoice Status")
                field_rows['category'] = pt4.PivotFields("Category")
                field_rows['company_name'] = pt4.PivotFields("Company Name")

                # insert row fields
                field_rows['invoice_status'].Orientation = 1
                field_rows['invoice_status'].Position = 1
                field_rows['invoice_status'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['category'].Orientation = 1
                field_rows['category'].Position = 2
                field_rows['category'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['company_name'].Orientation = 1
                field_rows['company_name'].Position = 3
                field_rows['company_name'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                        
                # Adding Invoice Due Date to Pivot Field Column
                field_columns = {}
                field_columns['due_date'] = pt4.PivotFields("Invoice Due Date")
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 1


                # Grouping column fields for Month, Quarter, Year
                field_columns['due_date'].AutoGroup()
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 2
                field_columns['due_date'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_columns['due_year'] = pt4.PivotFields("Years")
                field_columns['due_year'].Orientation = 2
                field_columns['due_year'].Position = 1
                field_columns['due_year'].RepeatLabels = 2
                field_columns['due_year'].ShowDetail = 1
                field_columns['due_year'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Setting Pivot Field "Quarter" to be hidden
                field_columns['due_quarter'] = pt4.PivotFields("Quarters")
                field_columns['due_quarter'].Orientation = 0
                        
                field_values = {}
                field_values['amount'] = pt4.PivotFields("Invoice Amount")


                # insert value fields
                field_values['amount'].Orientation = 4
                field_values['amount'].Function = -4157
                
                # Insert filter fields
                field_filters= {}
                
                field_filters['entity_name'] = pt4.PivotFields('Entity Name')
                field_filters['entity_name'].Orientation = 3
                
                field_filters['payment_method'] = pt4.PivotFields('Payment Method')
                field_filters['payment_method'].Orientation = 3
                
                # This puts "Entity Name" field on top of "Payment Method"
                field_filters['entity_name'].Position = 2
                
                field_filters['entity_name'].CurrentPage = "MPL Brands Inc (Default)"
                field_filters['payment_method'].CurrentPage = "Not Onboarded"
                
                selectRange = ws_6.Range("A:EV")
                selectRange.Font.Name = "Century Gothic"
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            def Category_Patco1(pt5):
                field_rows = {}
                field_rows['invoice_status'] = pt5.PivotFields("Invoice Status")
                field_rows['category'] = pt5.PivotFields("Category")
                field_rows['company_name'] = pt5.PivotFields("Company Name")

                # insert row fields
                field_rows['invoice_status'].Orientation = 1
                field_rows['invoice_status'].Position = 1
                field_rows['invoice_status'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['category'].Orientation = 1
                field_rows['category'].Position = 2
                field_rows['category'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['company_name'].Orientation = 1
                field_rows['company_name'].Position = 3
                field_rows['company_name'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                        
                # Adding Invoice Due Date to Pivot Field Column
                field_columns = {}
                field_columns['due_date'] = pt5.PivotFields("Invoice Due Date")
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 1


                # Grouping column fields for Month, Quarter, Year
                field_columns['due_date'].AutoGroup()
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 2
                field_columns['due_date'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_columns['due_year'] = pt5.PivotFields("Years")
                field_columns['due_year'].Orientation = 2
                field_columns['due_year'].Position = 1
                field_columns['due_year'].RepeatLabels = 2
                field_columns['due_year'].ShowDetail = 1
                field_columns['due_year'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Setting Pivot Field "Quarter" to be hidden
                field_columns['due_quarter'] = pt5.PivotFields("Quarters")
                field_columns['due_quarter'].Orientation = 0
                        
                field_values = {}
                field_values['amount'] = pt5.PivotFields("Invoice Amount")


                # insert value fields
                field_values['amount'].Orientation = 4
                field_values['amount'].Function = -4157
                
                # Insert filter fields
                field_filters= {}
                
                field_filters['entity_name'] = pt5.PivotFields('Entity Name')
                field_filters['entity_name'].Orientation = 3
                
                field_filters['payment_method'] = pt5.PivotFields('Payment Method')
                field_filters['payment_method'].Orientation = 3
                
                # This puts "Entity Name" field on top of "Payment Method"
                field_filters['entity_name'].Position = 2
                
                field_filters['entity_name'].CurrentPage = "Patco Brands"
                field_filters['payment_method'].CurrentPage = "Onboarded"
                
                selectRange = ws_7.Range("A:EV")
                selectRange.Font.Name = "Century Gothic"
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            def Category_Patco2(pt6):
                field_rows = {}
                field_rows['invoice_status'] = pt6.PivotFields("Invoice Status")
                field_rows['category'] = pt6.PivotFields("Category")
                field_rows['company_name'] = pt6.PivotFields("Company Name")

                # insert row fields
                field_rows['invoice_status'].Orientation = 1
                field_rows['invoice_status'].Position = 1
                field_rows['invoice_status'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['category'].Orientation = 1
                field_rows['category'].Position = 2
                field_rows['category'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_rows['company_name'].Orientation = 1
                field_rows['company_name'].Position = 3
                field_rows['company_name'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                        
                # Adding Invoice Due Date to Pivot Field Column
                field_columns = {}
                field_columns['due_date'] = pt6.PivotFields("Invoice Due Date")
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 1


                # Grouping column fields for Month, Quarter, Year
                field_columns['due_date'].AutoGroup()
                field_columns['due_date'].Orientation = 2
                field_columns['due_date'].Position = 2
                field_columns['due_date'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                field_columns['due_year'] = pt6.PivotFields("Years")
                field_columns['due_year'].Orientation = 2
                field_columns['due_year'].Position = 1
                field_columns['due_year'].RepeatLabels = 2
                field_columns['due_year'].ShowDetail = 1
                field_columns['due_year'].Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

                # Setting Pivot Field "Quarter" to be hidden
                field_columns['due_quarter'] = pt6.PivotFields("Quarters")
                field_columns['due_quarter'].Orientation = 0
                        
                field_values = {}
                field_values['amount'] = pt6.PivotFields("Invoice Amount")


                # insert value fields
                field_values['amount'].Orientation = 4
                field_values['amount'].Function = -4157
                
                # Insert filter fields
                field_filters= {}
                
                field_filters['entity_name'] = pt6.PivotFields('Entity Name')
                field_filters['entity_name'].Orientation = 3
                
                field_filters['payment_method'] = pt6.PivotFields('Payment Method')
                field_filters['payment_method'].Orientation = 3
                
                # This puts "Entity Name" field on top of "Payment Method"
                field_filters['entity_name'].Position = 2
                
                field_filters['entity_name'].CurrentPage = "Patco Brands"
                field_filters['payment_method'].CurrentPage = "Not Onboarded"
                
                selectRange = ws_7.Range("A:EV")
                selectRange.Font.Name = "Century Gothic"
                selectRange.Font.Size = 12
                selectRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                
            ws_5.Visible = False
            ws_1.Visible = False

            AP_PivotTable(pt)
            
            # Creating Count_PivotTable sheet
            Count_PivotTable(pt2)
            Not_Onboarded(ws_5,ws_4)
            
            # Creating Category_MPL sheet pivot table 1 on the left
            Category_MPL1(pt3)
            
            # Putting the position of the 2nd pivot table of Category_MPL sheet on the right starting at W1
            pt4 = pt_cache.CreatePivotTable(ws_6.Range("W1"), "MPL Brands - Not Onboarded")
            pt4.ColumnGrand = True
            pt4.RowGrand = True    
            Category_MPL2(pt4)
            
            Category_Patco1(pt5)
            
            # Putting the position of the 2nd pivot table of Category_Patco sheet on the right starting at W1
            pt6 = pt_cache.CreatePivotTable(ws_7.Range("W1"), "Patco Brands - Not Onboarded")
            pt6.ColumnGrand = True
            pt6.RowGrand = True
            Category_Patco2(pt6)
            
            # Set table style
            pt.TableStyle2 = "PivotStyleMedium9"
            pt2.TableStyle2 = "PivotStyleMedium9"
            pt3.TableStyle2 = "PivotStyleMedium9"
            pt4.TableStyle2 = "PivotStyleMedium9"
            pt5.TableStyle2 = "PivotStyleMedium9"
            pt6.TableStyle2 = "PivotStyleMedium9"
            
            self.print_success()
        except Exception as e:
            self.print_error(e)

mw = MainWindow()

app.exec_()

