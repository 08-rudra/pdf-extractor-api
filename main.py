from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from typing import List
import fitz  # PyMuPDF
import re
import logging

# Setup basic logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- PDF Extraction Logic ---
def extract_from_pdf(pdf_content: bytes, labels_from_excel: List[str]):
    # This dictionary will hold the final extracted data
    extracted_data = {label: "" for label in labels_from_excel}
    pdf_values = {}
    
    try:
        doc = fitz.open(stream=pdf_content, filetype="pdf")
        full_text = "".join(page.get_text("text") for page in doc)
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]

        # --- Single Pass Extraction ---
        for i, line in enumerate(lines):
            line_lower = line.lower()
            line_above = lines[i-1] if i > 0 else ""

            if "applicant name:" in line_lower and "co-applicant" not in line_lower:
                pdf_values["applicant_name"] = line.split(':')[1].strip()
            elif "year built:" in line_lower:
                pdf_values["year_built"] = line_above
            elif "heating last year updated:" in line_lower:
                pdf_values["heating_year"] = line_above
            elif "trampoline (y/n):" in line_lower:
                 if i + 1 < len(lines):
                    pdf_values["trampoline"] = lines[i+1]
            elif "slide (y/n):" in line_lower:
                pdf_values["slide"] = "No"
                if "☑" in line_lower: pdf_values["slide"] = "Yes"
            elif "pool:" in line_lower:
                pdf_values["pool"] = "No"
                if "☑" in line_lower: pdf_values["pool"] = "Yes"
        
        logger.info(f"Found values in PDF: {pdf_values}")

        # --- Map PDF Values to Simplified Excel Labels ---
        for label in labels_from_excel:
            label_lower = label.lower()
            if "applicant name" in label_lower:
                extracted_data[label] = pdf_values.get("applicant_name", "Not Found")
            elif "year built" in label_lower:
                extracted_data[label] = pdf_values.get("year_built", "Not Found")
            elif "heating" in label_lower:
                extracted_data[label] = pdf_values.get("heating_year", "Not Found")
            elif "trampoline" in label_lower:
                extracted_data[label] = pdf_values.get("trampoline", "Not Found")
            elif "slide" in label_lower:
                extracted_data[label] = pdf_values.get("slide", "Not Found")
            elif "pool" in label_lower:
                extracted_data[label] = pdf_values.get("pool", "Not Found")

    except Exception as e:
        logger.error(f"Error during PDF processing: {e}")
        raise HTTPException(status_code=500, detail=f"PDF processing failed: {e}")

    return extracted_data

# --- FastAPI Application ---
app = FastAPI(
    title="PDF Extractor API",
    description="Extracts data from Homeowners Insurance Questionnaires.",
    version="1.0.0"
)

@app.post("/extract-data/")
async def create_upload_file(labels: List[str] = Form(...), file: UploadFile = File(...)):
    """
    Accepts a PDF and a list of labels, returns extracted data as JSON.
    """
    if file.content_type != "application/pdf":
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a PDF.")
    
    pdf_content = await file.read()
    data = extract_from_pdf(pdf_content, labels)
    return data
```

**4. Create `requirements.txt`**
This file lists the project's dependencies, which our hosting service will use to install them.
```bash
pip freeze > requirements.txt
```

**5. Commit Your Backend Code**
Save the new files to your Git repository.
```bash
git add main.py requirements.txt
git commit -m "Feat: Implement FastAPI backend for PDF extraction"
```

---

### **Part 3: Deployment**

Now we put our API on the internet.

**1. Create a GitHub Repository**
Go to [GitHub.com](https://github.com), create a new repository (e.g., `pdf-extractor-api`), and follow the instructions to "push an existing repository from the command line." It will look something like this:
```bash
git remote add origin https://github.com/YourUsername/pdf-extractor-api.git
git branch -M main
git push -u origin main
```

**2. Deploy on Render**
**Render.com** is a modern, free-tier cloud host that is perfect for this.

1.  Sign up for a free account at Render.com.
2.  On your dashboard, click **New > Web Service**.
3.  Connect your GitHub account and select your `pdf-extractor-api` repository.
4.  Configure the service:
    * **Name:** `pdf-extractor-api` (or anything you like)
    * **Environment:** `Python 3`
    * **Build Command:** `pip install -r requirements.txt`
    * **Start Command:** `uvicorn main:app --host 0.0.0.0 --port 10000`
5.  Click **Create Web Service**.

Render will now build and deploy your application. After a few minutes, it will be live, and you will get a public URL like `https://pdf-extractor-api.onrender.com`. **This is your API URL.**


---

### **Part 4: The Final Excel Frontend**

The last step is to update the Excel macro to call your live API URL.

**Action**: Replace the macro in your Excel file with this final version. Remember to **paste your live API URL** where indicated.

```vb
' Add a reference to "Microsoft XML, v6.0" via Tools > References

' --- IMPORTANT: PASTE YOUR LIVE API URL FROM RENDER HERE ---
Private Const API_URL As String = "https://pdf-extractor-api.onrender.com/extract-data/"

Sub ExtractDataViaAPI()
    Dim sheet As Worksheet, lastRow As Long, i As Long
    Set sheet = ThisWorkbook.Sheets(1)
    
    ' --- 1. Get PDF File ---
    Dim fDialog As FileDialog, pdfPath As String
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.Title = "Select the PDF Questionnaire"
    If fDialog.Show <> -1 Then
        pdfPath = fDialog.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    ' --- 2. Get SIMPLIFIED Labels from Excel ---
    lastRow = sheet.Cells(sheet.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub ' No labels to process
    Dim labels() As String: ReDim labels(1 To lastRow - 1)
    For i = 2 To lastRow
        labels(i - 1) = sheet.Cells(i, 1).Value
    Next i
    
    ' --- 3. Build and Send the API Request ---
    Dim http As Object, boundary As String
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW" ' A standard boundary
    
    Dim reqBody As Object
    Set reqBody = CreateObject("ADODB.Stream")
    reqBody.Type = 1 ' adTypeBinary
    reqBody.Open
    
    ' Add labels
    For i = LBound(labels) To UBound(labels)
        reqBody.WriteText "--" & boundary & vbCrLf, 1
        reqBody.WriteText "Content-Disposition: form-data; name=""labels""" & vbCrLf & vbCrLf, 1
        reqBody.WriteText labels(i) & vbCrLf, 1
    Next i
    
    ' Add file
    reqBody.WriteText "--" & boundary & vbCrLf, 1
    reqBody.WriteText "Content-Disposition: form-data; name=""file""; filename=""" & Mid(pdfPath, InStrRev(pdfPath, "\") + 1) & """" & vbCrLf, 1
    reqBody.WriteText "Content-Type: application/pdf" & vbCrLf & vbCrLf, 1
    reqBody.Write CreateObject("ADODB.Stream").Open.LoadFromFile(pdfPath).Read
    reqBody.WriteText vbCrLf, 1
    reqBody.WriteText "--" & boundary & "--" & vbCrLf, 1
    
    ' --- 4. Send Request and Handle Response ---
    On Error GoTo HttpError
    http.Open "POST", API_URL, False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.send reqBody.Read
    
    If http.Status = 200 Then
        ' --- 5. Parse JSON Response (Requires a JSON Parser library for robust parsing) ---
        ' This is a simple parser. For complex JSON, a library is better.
        Dim jsonResponse As String: jsonResponse = http.responseText
        For i = 2 To lastRow
            Dim key As String: key = sheet.Cells(i, 1).Value
            If InStr(1, jsonResponse, """" & key & """:""") > 0 Then
                sheet.Cells(i, 2).Value = Split(Split(jsonResponse, """" & key & """:""")(1), """")(0)
            End If
        Next i
        MsgBox "Data extracted successfully!", vbInformation
    Else
        MsgBox "API Error " & http.Status & ": " & http.statusText & vbCrLf & vbCrLf & http.responseText, vbCritical
    End If
    Exit Sub

HttpError:
    MsgBox "An error occurred while contacting the server: " & Err.Description, vbCritical
End Sub
