import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Drive to PPT Generator", layout="centered")
st.title("📊 Professional PPT Generator (Recursive Subfolders)")

# -------------------------
# User Inputs
# -------------------------
campaign_name = st.text_input("📌 Campaign Name")
drive_link = st.text_input("🔗 Google Drive Folder Link")
generate_btn = st.button("🚀 Generate Presentation")

# -------------------------
# Google Drive Authentication
# -------------------------
def authenticate_drive():
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gdrive"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=credentials)

def extract_folder_id(link):
    if "folders/" not in link:
        st.error("Please provide a valid Google Drive Folder link")
        st.stop()
    return link.split("folders/")[1].split("?")[0]

def get_images_recursive(service, folder_id):
    """Recursively fetch all images from folder and subfolders"""
    all_images = []

    # Get images in current folder
    query = f"'{folder_id}' in parents and mimeType contains 'image/'"
    results = service.files().list(q=query, pageSize=1000).execute()
    images = results.get("files", [])
    all_images.extend(images)

    # Get subfolders
    folder_query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
    folders = service.files().list(q=folder_query, pageSize=1000).execute().get("files", [])

    for f in folders:
        all_images.extend(get_images_recursive(service, f["id"]))

    return all_images

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
# Generate PPT Logic
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
        images = get_images_recursive(service, folder_id)

        if not images:
            st.error("No images found in the folder or its subfolders.")
            st.stop()

        prs = Presentation()

        # Title slide
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        slide.shapes.title.text = campaign_name
        slide.placeholders[1].text = "Auto Generated Presentation"

        # Slide settings
        img_per_slide = 4
        col_positions = [Inches(0.5), Inches(5)]  # 2 columns
        row_positions = [Inches(1.5), Inches(4)]  # 2 rows
        width = Inches(4.0)
        height = Inches(2.5)

        # Background color
        bg_color = RGBColor(230, 230, 250)  # light lavender

        # Generate slides
        for i in range(0, len(images), img_per_slide):
            slide_layout = prs.slide_layouts[6]  # blank slide
            slide = prs.slides.add_slide(slide_layout)

            # Set background color
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = bg_color

            slide_images = images[i:i+img_per_slide]
            footer_names = []

            for idx, img in enumerate(slide_images):
                image_stream = download_image(service, img["id"])
                col = idx % 2
                row = idx // 2
                left = col_positions[col]
                top = row_positions[row]
                slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
                footer_names.append(img["name"])

            # Footer text with all image names on this slide
            footer_text = " | ".join(footer_names)
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(7.7), Inches(9), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = footer_text
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(50, 50, 50)

        # Save PPT to memory
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success("Professional Presentation Generated Successfully!")
        st.download_button(
            label="📥 Download PPT",
            data=ppt_io,
            file_name=f"{campaign_name}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        st.error(f"Error: {e}")
