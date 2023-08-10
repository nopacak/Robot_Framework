from RPA.Tables import Tables
from robocorp.tasks import task
from robocorp import browser, http
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive


page = browser.page()

def open_the_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads csv file from the given URL"""
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def close_POPUP():
    """Closing pop-up message"""
    page.click("text=OK")

def fill_order_form(order):
    """Fills in the order form, clicks the 'Order' button"""
    body_value = str(order["Body"])
    close_POPUP()
    page.select_option("select#head", value= str(order["Head"]))
    page.click(f"#id-body-{body_value}")
    page.fill("//input[@placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", str(order["Address"]))
    page.click("text=Preview")
    while page.is_visible("#order-completion") == False:
        page.click("#order")

def fill_form_with_csv_data():
    """Read data from csv and fill in the order form"""
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])
    for row in orders:
        fill_order_form(row)
        export_as_pdf(row)
        order_another_robot()

def order_another_robot():
    """Click the 'Order another robot' button"""
    page.click("text=Order another robot")

def export_as_pdf(order):
    """Export the data as a PDF"""
    order_no = str(order["Order number"])
    order_summary_html = page.locator("#receipt.alert.alert-success").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(order_summary_html, f"temp/order_receipt_{order_no}.pdf")
    page.screenshot(path=f"temp/robot{order_no}.png")
    pdf.add_watermark_image_to_pdf(
        image_path=f"temp/robot{order_no}.png",
        source_path=f"temp/order_receipt_{order_no}.pdf",
        output_path=f"to_be_zipped/order_{order_no}.pdf")

def create_zip():
    """Saving all PDF order receipts in a ZIP directory"""
    Archive().archive_folder_with_zip("to_be_zipped", "Archived_orders.zip", recursive=True)
    
def cleanup():
    """Deleting unnecessary directories"""
    FileSystem().remove_directory("temp", recursive=True)
    FileSystem().remove_directory("to_be_zipped", recursive=True)





@task
def robot_orders_python():
    """Insert the orders, export them as a PDF and ZIP them"""
    browser.configure(slowmo=100)
    open_the_order_website()
    download_csv_file()
    fill_form_with_csv_data()
    create_zip()
    cleanup()




