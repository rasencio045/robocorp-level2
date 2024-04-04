from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from random import randrange
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=50,
    )

    download_order_csv()
    open_browser()
    fill_form_with_excel_data()

def download_order_csv():
    """download file"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def open_browser():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    csv = Tables()
    orders = csv.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Address"])
    for row in orders:
        fill_and_submit_order_form(row)
    zipFile = Archive()
    zipFile.archive_folder_with_zip("output/pdf/", "output.zip", include="*.pdf")
    

def fill_and_submit_order_form(row):
    page = browser.page()
    page.click("text=OK")
    page.select_option("#head", str(row["Head"]))
    page.click("id=id-body-"+row["Body"])
    page.get_by_placeholder("Enter the part number for the legs").fill(str(randrange(1,6)))
    page.fill("#address", str(row["Address"]))
    page.click("id=preview")
    page.locator("button").get_by_text("ORDER").click()
    if page.get_by_text("-ROBO-ORDER-").check():
        print("works")
    else:
        print("Server error")
    order_number = page.get_by_text("-ROBO-ORDER-").inner_html()
    save_order(order_number)

def save_order(order_number):
    page = browser.page()
    pdf = PDF()
    image_path="output/pdf/" + order_number + ".png"
    pdf_path="output/pdf/" + order_number + ".pdf"
    page.screenshot(path=image_path)
    order_detail = page.locator("id=receipt").inner_html()
    image = ("<img src=\"" + image_path + "\" width=\"600\"/>")
    pdf.html_to_pdf([order_detail, image], pdf_path)
    
    page.locator("button").get_by_text("ORDER").click()