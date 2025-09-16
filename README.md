PDF Extractor API
This project is a web API built with Python and FastAPI to extract structured data from PDF questionnaire forms. An Excel macro sends a PDF to the API, which then processes the document and returns the extracted data in JSON format.

Features
Extracts data from a specific PDF form layout.

Provides a simple API endpoint for integration.

Centralized logic for easy updates.

Setup & Deployment
Clone the repository.

Create a Python virtual environment and install dependencies from requirements.txt.

Deploy to a cloud hosting service like Render.


**6. Make Your First Commit**
Now, let's save our initial project structure to Git.
```bash
git add .
git commit -m "Initial commit: Set up project structure and documentation"

Part 2: Building the Python API Backend
Now we'll create the Python application that will live on the server.

1. Create a Virtual Environment
This isolates your project's dependencies.

python -m venv venv

Activate it:

Windows: venv\Scripts\activate

macOS/Linux: source venv/bin/activate

2. Install Dependencies

pip install "fastapi[all]" python-multipart PyMuPDF
