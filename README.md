# RRT---Report-Generator-
Streamlit Report Generator App
This Streamlit app reads an Excel workbook with multiple sheets and generates printable reports for each sheet. It processes pivot tables and associated graphs, then outputs professional .docx reports saved in the current working directory.

🚀 Features
📂 Upload a multi-sheet Excel file

📊 Automatically reads pivot tables and charts from each sheet

📝 Generates a clean .docx report for each sheet

💾 Saves the reports directly to your current folder

🛠️ Installation
Clone this repository:

bash
Copy
Edit
git clone https://github.com/your-username/streamlit-report-generator.git
cd streamlit-report-generator
Create a virtual environment (optional but recommended):

bash
Copy
Edit
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
Install dependencies:

bash
Copy
Edit
pip install -r requirements.txt
🏃 Usage
Launch the Streamlit app:

bash
Copy
Edit
streamlit run app.py
Upload your Excel workbook when prompted.

Reports for each sheet will be saved in the directory as .docx files.

📁 File Structure
bash
Copy
Edit
├── app.py               # Main Streamlit app
├── requirements.txt     # Python dependencies
└── README.md            # This file
📌 Notes
Excel file must contain pivot tables and graphs in the expected format.

Reports are generated using python-docx and include visualizations.

📄 License
MIT License. Free to use and modify.
