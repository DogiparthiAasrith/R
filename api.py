# api.py
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import pandas as pd
import io
import uvicorn

# --- Copy your existing utility and processing functions here ---
# num, safe_int, safeval, find_header_row, read_bs_and_pl, process_financials
# Make sure to copy all of them to this file.
# ---------------------------------------------------------------

# Create the FastAPI app
app = FastAPI(title="Financial Mapping API", description="An API to process financial statements and generate reports.")

@app.get("/")
def read_root():
    return {"message": "Welcome to the Financial Mapping API! Use the /upload-file endpoint to process your data."}

@app.post("/upload-file/")
async def process_excel_file(file: UploadFile = File(...)):
    """
    Accepts an Excel file, processes it, and returns the generated financial reports as JSON.
    """
    if not file.filename.endswith(('.xls', '.xlsx')):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .xls or .xlsx file.")
        
    try:
        # Read the uploaded file into a BytesIO object
        input_file = io.BytesIO(await file.read())

        # Call your existing processing functions
        bs_df, pl_df = read_bs_and_pl(input_file)
        bs_out, pl_out, notes, totals = process_financials(bs_df, pl_df)

        # Convert DataFrames to JSON-serializable dictionaries
        # You need to convert all notes DataFrames as well
        notes_dict = {}
        for label, df in notes:
            notes_dict[label] = df.to_dict(orient='split')

        response_data = {
            "bs_out": bs_out.to_dict(orient='split'),
            "pl_out": pl_out.to_dict(orient='split'),
            "notes": notes_dict,
            "totals": totals
        }
        
        return JSONResponse(content=response_data, status_code=200)

    except Exception as e:
        print(f"Error processing file: {e}")
        raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")

# To run this API locally, you would execute: uvicorn api:app --reload
if __name__ == "__main__":
    uvicorn.run("api:app", host="127.0.0.1", port=8000, reload=True)