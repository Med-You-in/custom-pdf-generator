## **Custom PDF Generator**

### **Project Documentation**

**Version**: 1.0
**Authors**: Med-You-in

---

## **Table of Contents**

1. [Executive Summary](#executive-summary)
2. [System Architecture](#system-architecture)
3. [Features](#features)
4. [Installation and Setup](#installation-and-setup)
5. [Usage](#usage)
6. [Contributing](#contributing)
7. [License](#license)
8. [Contact](#contact)

---

## **1. Executive Summary**

The Custom PDF Generator is a tool designed to streamline the creation of custom PDFs using predefined templates and data inputs. It focuses on improving efficiency and providing a user-friendly interface for generating and managing PDFs.

Key objectives:

- **Objective 1**: Provide an intuitive interface for generating custom PDFs.
- **Objective 2**: Allow seamless integration of data from various sources.
- **Objective 3**: Enable easy customization of PDF templates.

Target Users:

- **Developers**: Integrate PDF generation into applications.
- **Designers**: Create and customize PDF templates.
- **End Users**: Generate PDFs using predefined templates.

---

## **2. System Architecture**

### **Technology Stack**

- **Backend**: Python
- **Frontend**: HTML, CSS
- **Libraries**: Pandas, Jinja2, pdfkit, openpyxl
- **Tools**: wkhtmltopdf

### **Core Components**

- `templates/`: Contains PDF templates.
- `data/`: Manages data inputs for PDF generation.
- `scripts/`: Handles PDF generation logic.
- `config/`: Manages configuration settings.
- `hooks/`: Contains Git hooks for repository management.

---

## **3. Features**

### **Core Functionalities**

1. **Template-Based Generation**: Create PDFs using predefined templates.
2. **Data Integration**: Seamlessly integrate data from Excel files into your PDFs.
3. **Customization**: Easily customize the appearance and content of your PDFs.
4. **Automated Workflows**: Use Git hooks to automate tasks like pre-commit checks and updates.

---

## **4. Installation and Setup**

### **Prerequisites**

- Git installed on your machine.
- Python installed.
- `wkhtmltopdf` installed.

### **Setup Instructions**

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/Med-You-in/custom-pdf-generator.git
   cd custom-pdf-generator
   ```

2. **Install Required Python Packages**:

   ```bash
   pip install -r requirements.txt
   ```

3. **Install `wkhtmltopdf`**:

   - Download from [wkhtmltopdf.org](https://wkhtmltopdf.org/downloads.html).
   - Follow the installation instructions for your operating system.
   - Add `wkhtmltopdf` to your system's PATH.

4. **Set Up the Project Folder**:

   - Create a folder structure as follows:
     ```
     custom-pdf-generator
     ├── .env
     ├── .env.example
     ├── .gitignore
     ├── assets
     │   ├── 0_image.png
     │   ├── 1_image.png
     │   ├── 2_image.png
     │   └── styles.css
     ├── data
     │   └── data-1.xlsx
     ├── documents
     │   └── generated_pdfs
     ├── README.md
     ├── requirements.txt
     ├── scripts
     │   ├── convToPdf-manual.py
     │   ├── convToPdf.py
     │   └── image_downloader.py
     └── templates
         └── pdf_template.html
     ```

5. **Prepare the Excel File**:

   - Create an Excel file named `data-1.xlsx` in the `data/` folder.
   - Add the following columns:
     - `Column1`: Text for the first column in the template.
     - `Column2`: Text for the second column in the template.
     - `Column3`: Text for the third column in the template.
     - `image`: Path to the image file (e.g., `path/to/image1.jpg`).

6. **Prepare the HTML Template**:

   - Update the `pdf_template.html` file in the `templates/` folder.

### **Troubleshooting**

- **`wkhtmltopdf` not found**: Ensure the path to `wkhtmltopdf` is correct in the script.
- **Missing dependencies**: Run `pip install -r requirements.txt` again.
- **Image not found**: Verify the `image` column in the Excel file contains valid paths to image files.

---

## **5. Usage**

### **Basic Workflow**

1. **Select Template**: Choose a predefined template.
2. **Input Data**: Enter the data to be integrated into the PDF.
3. **Generate PDF**: Click generate to create the PDF.

### **Examples**

- **Scenario A**:
  - Input: User data for a report.
  - Output: Customized PDF report.

### **Scripts Overview**

#### **convToPdf-manual.py**

- **Purpose**: Manually generate a PDF using predefined data.
- **Usage**:
  ```bash
  python scripts/convToPdf-manual.py
  ```
- **Description**: This script manually inserts data into the template and generates a PDF. It's useful for testing the template and ensuring everything is set up correctly.

#### **convToPdf.py**

- **Purpose**: Generate PDFs from data in an Excel file.
- **Usage**:
  ```bash
  python scripts/convToPdf.py
  ```
- **Description**: This script reads data from an Excel file, cleans it, and generates PDFs using the specified template. It logs any exceptions and saves them to an Excel file.

#### **image_downloader.py**

- **Purpose**: Download images from URLs listed in an Excel file.
- **Usage**:
  ```bash
  python scripts/image_downloader.py
  ```
- **Description**: This script reads URLs from an Excel file, downloads the images, and saves them to a specified folder. It logs the download process and any errors encountered.

---

## **6. Contributing**

We welcome contributions from the community! If you'd like to contribute, please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Make your changes and commit them with descriptive messages.
4. Push your branch to your fork.
5. Create a pull request to the main repository.

---

## **7. License**

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## **8. Contact**

For any questions or issues, please open an issue on the GitHub repository or contact the project maintainers.

---

Thank you for using the Custom PDF Generator! We hope it meets your needs and look forward to your contributions.
