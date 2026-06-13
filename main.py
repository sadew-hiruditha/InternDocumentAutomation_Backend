from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from docx import Document
import pandas as pd
import io
import zipfile

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Manual Input සඳහා ආකෘතිය


class InternData(BaseModel):
    name: str
    nic: str
    address: str
    start_date: str
    end_date: str

# දිනවල අගින් එන 00:00:00 කපා ඉවත් කිරීමේ function එක


def clean_date(date_val):
    if pd.isna(date_val) or not str(date_val).strip():
        return ""
    return str(date_val).split(" ")[0]

# දත්ත එක සමාන කිරීම (Excel වලින් ආවත්, Form එකෙන් ආවත්)


def normalize_data(raw_data):
    return {
        "name": str(raw_data.get('Name', raw_data.get('name', ''))).strip(),
        "nic": str(raw_data.get('NIC', raw_data.get('nic', ''))).strip(),
        "address": str(raw_data.get('Address', raw_data.get('address', ''))).strip(),
        "start_date": clean_date(raw_data.get('Start_Date', raw_data.get('start_date', ''))),
        "end_date": clean_date(raw_data.get('End_Date', raw_data.get('end_date', '')))
    }


def generate_ol_bytes(data):
    doc = Document()
    doc.add_paragraph(f"{data['start_date']}\n")
    doc.add_paragraph(f"{data['name']}\n{data['address']}\n")

    first_name = data['name'].split()[0] if data['name'] else 'Intern'
    doc.add_paragraph(f"Dear {first_name},\n")

    doc.add_paragraph(
        f"We are pleased to offer you a period of internship in the above company from {data['start_date']} to {data['end_date']}. "
        "We expect you to make use of this period to familiarize yourself with the corporate world by participating in our day to day operations along with our employees."
    )
    doc.add_paragraph(
        "You should liaise with the undersigned in relation to all matters during this period.\n")
    doc.add_paragraph(
        "Yours faithfully,\nCeylon Cold Store Plc\n\n\nWasantha mudalige\nHead of The Human Resource Operation\n")

    doc.add_paragraph("_" * 50 + "\n")
    doc.add_paragraph(
        f"I am pleased to accept this offer of 06 months internship commencing {data['start_date']} on the basis given above.\n")
    doc.add_paragraph(
        "Signature: _______________________      Date: _______________________\n")
    doc.add_paragraph(f"Name: {data['name']}      NIC number: {data['nic']}")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream.getvalue()


def generate_nda_bytes(data):
    doc = Document()
    doc.add_heading("NON-DISCLOSURE AGREEMENT", level=1)
    doc.add_paragraph(f"\nThis Agreement Made On This {data['start_date']}\n")
    doc.add_paragraph(
        f"Between Mr/Ms. {data['name']} (Holder Of National Identity Card Bearing The Number {data['nic']}) "
        f"Of {data['address']} (Hereinafter Referred To As The 'First Party')\n"
    )
    doc.add_paragraph(
        "And Whereas The First Part Is Desires Of Outsourced Intern In The Ceylon Cold Stores PLC, "
        "In Human Resource Department Of The Second Party And The Second Party Has Agreed To Such Outsourced Contract."
    )
    doc.add_paragraph(
        "\n\n_______________________\nSignature of the First Party\n")
    doc.add_paragraph(
        "\nAuthorized signature of the second party (Ceylon cold store Plc.,)\n\nWasantha mudalige\nHead of the human resource operation")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream.getvalue()

# ZIP ෆයිල් එක හදන පොදු function එක (Folders 2ක වෙනම දානවා)


def create_zip_archive(data_list):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for raw_data in data_list:
            data = normalize_data(raw_data)
            if not data['name'] or data['name'] == 'nan':
                continue

            safe_name = data['name'].replace(" ", "_")
            ol_file = generate_ol_bytes(data)
            nda_file = generate_nda_bytes(data)

            # වෙනම Folders වලට දාන කොටස
            zip_file.writestr(
                f"Offer_Letters/{safe_name}_Offer_Letter.docx", ol_file)
            zip_file.writestr(f"NDAs/{safe_name}_NDA.docx", nda_file)

    zip_buffer.seek(0)
    return zip_buffer

# 1. Manual Form Submit එකට


@app.post("/generate")
async def generate_single(data: InternData):
    try:
        zip_buffer = create_zip_archive([data.dict()])
        return StreamingResponse(
            zip_buffer,
            media_type="application/x-zip-compressed",
            headers={
                "Content-Disposition": f"attachment; filename=Intern_Docs_{data.name.replace(' ', '_')}.zip"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 2. Bulk Excel Upload එකට


@app.post("/generate-bulk")
async def generate_bulk(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(
            status_code=400, detail="කරුණාකර Excel (.xlsx) ෆයිල් එකක් පමණක් Upload කරන්න.")

    try:
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))

        # DataFrame එක dictionary list එකක් බවට පත් කිරීම
        data_list = df.to_dict(orient='records')

        zip_buffer = create_zip_archive(data_list)
        return StreamingResponse(
            zip_buffer,
            media_type="application/x-zip-compressed",
            headers={
                "Content-Disposition": "attachment; filename=Bulk_Intern_Docs.zip"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

        if __name__ == "__main__":
            import uvicorn
            uvicorn.run(app, host="0.0.0.0", port=8080)
