# Restaurant Expiry Guard

**A Smart Restaurant Inventory and Expiry Tracking System**

## Project Overview

Restaurant Expiry Guard is a Python-based inventory management application developed using Streamlit. The project aims to simplify restaurant inventory management by helping users track ingredient quantities, monitor expiry dates, manage stock usage, and reduce unnecessary food waste.

Restaurants handle large quantities of perishable ingredients every day. Managing their expiry dates manually can be time-consuming and may lead to food wastage. This project provides a simple digital solution to monitor inventory and notify users about ingredients that require attention.

The application combines inventory tracking, image-based expiry date recognition, Excel data management, and email notifications into a single interface.

## Key Features

- **Inventory Dashboard:** Displays in-stock, expiring soon, expired, and closed inventory batches.
- **Stock Management:** Allows users to record new ingredient deliveries, quantities, suppliers, categories, and storage locations.
- **Expiry Date Recognition:** Uses Optical Character Recognition (OCR) to extract possible expiry dates from uploaded product label images.
- **Expiry Monitoring:** Automatically checks expiry dates and classifies ingredients based on configurable alert periods.
- **Stock Usage Tracking:** Records ingredients as used or discarded and updates the remaining quantities.
- **Email Notifications:** Sends email alerts for expired and soon-to-expire ingredients.
- **Excel Integration:** Stores inventory information and stock movement history in an Excel workbook.
- **Inventory Export:** Allows users to download inventory records in Excel format.
- **Customizable Alerts:** Allows users to configure how many days before expiry an ingredient should trigger a warning.

## Technologies Used

| Technology | Purpose |
|---|---|
| Python | Core application development |
| Streamlit | Interactive web application interface |
| Pandas | Data handling and inventory presentation |
| OpenPyXL | Excel workbook creation and management |
| Pytesseract | Optical Character Recognition |
| Pillow (PIL) | Image processing for OCR |
| Python-dotenv | Environment variable management |
| SMTP | Email notification functionality |

## Project Structure

```text
Restaurant-Expiry-Guard/
│
├── app.py
├── data/
│   └── inventory.xlsx
├── .env
├── .env.example
├── .gitignore
├── requirements.txt
└── README.md
```

The inventory Excel workbook is automatically created when the application runs for the first time.

## Installation and Setup

### 1. Clone the Repository

```bash
git clone <your-repository-url>
```

### 2. Navigate to the Project Directory

```bash
cd Restaurant-Expiry-Guard
```

### 3. Create a Virtual Environment

```bash
python -m venv venv
```

Activate the environment:

**Windows:**
```bash
venv\Scripts\activate
```

**macOS / Linux:**
```bash
source venv/bin/activate
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

### 5. Install Tesseract OCR

Install Tesseract OCR separately on your system to enable expiry date recognition from uploaded product images.

Ensure that Tesseract is accessible to Pytesseract.

OCR functionality is optional, and expiry dates can also be entered manually.

### 6. Run the Application

```bash
streamlit run app.py
```

The application will open in your default web browser.

## Application Workflow

1. **Add Inventory:** Enter ingredient details, received quantity, supplier information, and expiry dates.
2. **Verify Expiry Dates:** Optionally upload a product label image to extract possible expiry dates using OCR.
3. **Monitor Inventory:** View ingredient details and expiry statuses through the dashboard.
4. **Track Stock Usage:** Record quantities of ingredients that are used or discarded.
5. **Receive Notifications:** Send email alerts for ingredients approaching or past their labelled dates.
6. **Export Records:** Download inventory information in Excel format for record-keeping.

## Expiry Status Classification

| Status | Description |
|---|---|
| OK | Ingredient is within its valid date range. |
| EXPIRING SOON | Ingredient is approaching its expiry date. |
| EXPIRED | The use-by or expiry date has passed. |
| PAST BEST BEFORE | The best-before date has passed and requires quality review. |
| CLOSED | The batch has been completely used or discarded. |

## Email Notification Configuration

The application supports email notifications through SMTP.

Create a `.env` file in the project directory and configure the following variables:

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your_email@gmail.com
SMTP_PASSWORD=your_app_password
ALERT_TO=recipient@gmail.com
```

For Gmail, use an App Password where required.

To send notifications, open the Settings & Alerts section and select the Send Alert Email Now option.

The application also supports command-line expiry checks:

```bash
python app.py --check-alerts
```

This command can be scheduled to execute daily using an operating system task scheduler.

**Security:** Never upload your `.env` file or email credentials to GitHub.

## Technical Implementation

- **Inventory Storage:** Uses an Excel workbook with separate Inventory, Settings, and Movements worksheets.
- **Expiry Calculation:** Uses Python date handling to classify inventory according to expiry dates and configurable warning periods.
- **OCR Processing:** Processes uploaded images using Pillow and Pytesseract to identify possible expiry dates.
- **Data Validation:** Validates ingredient names, quantities, stock movements, and expiry date confirmations before saving records.
- **Data Integrity:** Uses temporary workbook saving and file replacement to reduce the risk of incomplete Excel writes.
- **Email Automation:** Uses Python's SMTP libraries to send notifications for inventory batches requiring attention.

## Key Learning Outcomes

Through this project, I gained practical experience in:

- Developing interactive applications using Python and Streamlit.
- Managing structured inventory data using Excel and OpenPyXL.
- Implementing image processing and OCR functionality.
- Working with date validation and expiry monitoring logic.
- Integrating email notifications using SMTP.
- Handling user input, validation, and file operations.
- Designing a practical software solution to address a real-world problem.

## Future Enhancements

- User authentication and role-based access.
- Cloud database integration for centralized inventory management.
- Advanced inventory analytics and visual reports.
- Automatic low-stock notifications.
- Improved OCR accuracy and expiry date recognition.
- Multi-restaurant inventory management.
- A more advanced reporting and management dashboard.

## Limitations

- Inventory information is currently stored locally in an Excel workbook.
- OCR-generated dates are suggestions and must be verified against the original product label.
- Email notifications require SMTP configuration.
- Daily notifications require an external scheduler; they do not run automatically when the application is closed.

## Author

**Prisha Shinde**

## License

This project is intended for educational and personal use. An appropriate open-source license can be added if the project is distributed publicly.
