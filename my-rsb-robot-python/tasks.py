from robocorp.tasks import task
from robocorp import browser, http, excel
from RPA.PDF import PDF


page = browser.page()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in(username, password):
    """Fills in the login form and clicks the 'Log in' button"""
    page.fill("#username", username)
    page.fill("#password", password)
    page.click("button:text('Log in')")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    workbook = excel.open_workbook("SalesData.xlsx")
    worksheet = workbook.worksheet("data").as_table(header=True)
    for row in worksheet:
        fill_and_submit_sales_form(row)

def collect_results():
    """Take a screenshot of the page"""
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the data as a PDF"""
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out():
    """Press the 'Log Out' button"""
    page.click("text=Log out")




@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(slowmo=100)
    open_the_intranet_website()
    log_in("maria", "thoushallnotpass")
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()