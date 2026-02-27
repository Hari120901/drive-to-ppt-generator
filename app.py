import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Drive to PPT Generator", layout="centered")

st.title("📊 Automated PPT Generator")

# -------------------------
# User Inputs
# -------------------------
campaign_name = st.text_input("📌 Campaign Name")
drive_link = st.text_input("🔗 Google Drive Folder Link")

generate_btn = st.button("🚀 Generate Presentation")

# -------------------------
# Authenticate Google Drive
# -------------------------
def authenticate_drive():
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gdrive"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=credentials)

def extract_folder_id(link):
    return link.split("folders/")[1].split("?")[0]

def get_subfolders(service, parent_id):
    query = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder'"
    results = service.files().list(q=query).execute()
    return results.get("files", [])

def get_images(service, folder_id):
    query = f"'{folder_id}' in parents and mimeType contains 'image/'"
    results = service.files().list(q=query).execute()
    return results.get("files", [])

def download_image(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# -------------------------
# Generate PPT
# -------------------------
if generate_btn:

    if not campaign_name:
        st.warning("Please enter Campaign Name")
        st.stop()

    if not drive_link:
        st.warning("Please enter Drive Folder Link")
        st.stop()

    try:
        service = authenticate_drive()
        folder_id = extract_folder_id(drive_link)
        subfolders = get_subfolders(service, folder_id)

        if not subfolders:
            st.error("No subfolders found.")
            st.stop()

        prs = Presentation()

        # Title slide
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = campaign_name
        slide.placeholders[1].text = "Auto Generated Presentation"

        for folder in subfolders:

            images = get_images(service, folder["id"])
            if not images:
                continue

            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = folder["name"]

            left_positions = [0.5, 3.5, 6.5]
            top_positions = [1.5, 4]

            img_count = 0

            for img in images:
                if img_count >= 6:
                    break

                image_stream = download_image(service, img["id"])

                left = Inches(left_positions[img_count % 3])
                top = Inches(top_positions[img_count // 3])

                slide.shapes.add_picture(
                    image_stream,
                    left,
                    top,
                    width=Inches(2.8)
                )

                img_count += 1

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success("Presentation Generated Successfully!")

        st.download_button(
            label="📥 Download PPT",
            data=ppt_io,
            file_name=f"{campaign_name}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        st.error(f"Error: {e}")
