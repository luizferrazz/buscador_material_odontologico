🦷 Dental Materials Scraper

This project automates the process of searching dental materials on the Dental Cremer
 website using Selenium, extracts product information (name, material, price, URL, promotions), and exports the data into an Excel file grouped by material.

 📌 Features

Reads a list of materials from materials.txt.

Uses Selenium WebDriver (headless Chrome) to search each material.

Extracts:

Product name
Material searched
Price
URL
Promotion (if available)

Saves results in found_materials.xlsx:
Each material has its own worksheet.
Includes headers and structured data.

⚙️ Requirements

Python 3.8+
Google Chrome installed
Dependencies (see below)

📦 Installation

Clone the repository and install dependencies:

git clone [repository](https://github.com/luizferrazz/buscador_material_odontologico)
cd buscador_material_odontologico
pip install -r requirements.txt

Contents of requirements.txt:

selenium
webdriver-manager
pandas
openpyxl

📝 Usage

Add the materials you want to search inside Materials/materials.txt (one per line).

alginate
dental mirror
composite resin


Run the scraper:

python scraper.py


Check the results in Materials/found_materials.xlsx.
Each material will have its own Excel sheet with all products found.

🧩 How It Works

Loads materials.txt.

Opens Dental Cremer with Selenium in headless mode.

Searches each material and scrapes product cards.

Collects structured data in Python.

Creates/updates an Excel workbook with grouped sheets per material.

⚠️ Notes

Site structure may change, breaking the scraper. Adjust XPaths/CSS selectors if needed.

Headless Chrome is used by default. Remove --headless in setup_webdriver() to watch the scraping in real-time.

The Excel sheet name is limited to 31 characters (Excel limitation).

📌 Example Output

Example sheet for "alginate":

Name	Material	Price	URL	Promotion
Alginate Fast Set	alginate	R$ 29.90	https://www.dentalcremer.com.br/...	No
Alginate Chromatic	alginate	R$ 35.50	https://www.dentalcremer.com.br/...	15% Off

📜 License

This project is licensed under the MIT License.
