import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Flipkart Minutes Style PPT Gen", layout="centered")
st.title("📊 Professional PPT Generator")

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

def get_folder_metadata(service, folder_id):
    """Fetch the actual name of the folder from Drive"""
    folder = service.files().get(fileId=folder_id, fields="name").execute()
    return folder.get("name", "Unknown Location")

def get_images_recursive(service, folder_id):
    all_images = []
    query = f"'{folder_id}' in parents and mimeType contains 'image/'"
    results = service.files().list(q=query, pageSize=1000).execute()
    all_images.extend(results.get("files", []))

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
# User Inputs
# -------------------------
campaign_name = st.text_input("📌 Campaign Name (e.g., Flipkart Minutes)")
drive_link = st.text_input("🔗 Google Drive Folder Link")
generate_btn = st.button("🚀 Generate Presentation")

if generate_btn:
    if not campaign_name or not drive_link:
        st.warning("Please fill in all fields")
        st.stop()

    try:
        service = authenticate_drive()
        folder_id = extract_folder_id(drive_link)
        location_name = get_folder_metadata(service, folder_id)
        images = get_images_recursive(service, folder_id)

        if not images:
            st.error("No images found.")
            st.stop()

        prs = Presentation()
        # Set slide dimensions to Widescreen (13.33 x 7.5) or Standard (10 x 7.5)
        # Using standard 10x7.5 based on your reference image aspect ratio
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)

        # Layout constants
        TEAL_COLOR = RGBColor(0, 140, 170) 
        
        for i in range(0, len(images), 3):
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout

            # 1. Create Teal Header Bar
            header_rect = slide.shapes.add_shape(
                1, # Rectangle
                Inches(0.2), Inches(0.2), Inches(5), Inches(0.8)
            )
            header_rect.fill.solid()
            header_rect.fill.fore_color.rgb = TEAL_COLOR
            header_rect.line.fill.background() # No border

            # 2. Add Location Text (Folder Name) inside Teal Bar
            loc_text = header_rect.text_frame
            loc_text.text = location_name
            p_loc = loc_text.paragraphs[0]
            p_loc.font.bold = True
            p_loc.font.size = Pt(20)
            p_loc.font.color.rgb = RGBColor(255, 255, 255)
            p_loc.alignment = PP_ALIGN.CENTER

            # 3. Add Advertiser Label (Campaign Name)
            adv_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6), Inches(0.5))
            p_adv = adv_box.text_frame.paragraphs[0]
            p_adv.text = f"Advertiser: {campaign_name}"
            p_adv.font.bold = True
            p_adv.font.size = Pt(24)
            p_adv.font.color.rgb = RGBColor(0, 0, 0)

            # 4. Place 3 Images Side-by-Side
            slide_images = images[i:i+3]
            img_width = Inches(3.0)
            img_height = Inches(4.5) # Tall portrait style
            start_left = Inches(0.3)
            gap = Inches(0.2)
            top_pos = Inches(2.2)

            for idx, img in enumerate(slide_images):
                img_stream = download_image(service, img["id"])
                left = start_left + (idx * (img_width + gap))
                
                # Add a black border/frame effect around image
                slide.shapes.add_picture(img_stream, left, top_pos, width=img_width, height=img_height)
                
                # Optional: Thin border around the image
                rect = slide.shapes.add_shape(1, left, top_pos, img_width, img_height)
                rect.fill.background()
                rect.line.color.rgb = RGBColor(0, 0, 0)
                rect.line.width = Pt(1.5)

        # Save and Download
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success(f"Generated slides for {location_name}!")
        st.download_button(
            label="📥 Download PPT",
            data=ppt_io,
            file_name=f"{location_name}_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
