import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Flipkart Minutes Style PPT Gen", layout="wide")
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

def get_subfolders(service, parent_id):
    query = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    return results.get("files", [])

def get_images_in_folder(service, folder_id):
    query = f"'{folder_id}' in parents and mimeType contains 'image/' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
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
# PPT Generation Logic
# -------------------------
def create_title_slide(prs, campaign_name):
    """Creates the first page based on the 'Proof of Play' screenshot"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Placeholder for the Teal graphic bar
    teal_bar = slide.shapes.add_shape(1, 0, Inches(2.2), Inches(3), Inches(1.5))
    teal_bar.fill.solid()
    teal_bar.fill.fore_color.rgb = RGBColor(0, 140, 170)
    teal_bar.line.fill.background()

    # Campaign Name (Placed above Proof of Play)
    camp_box = slide.shapes.add_textbox(Inches(0.2), Inches(5.0), Inches(5), Inches(0.5))
    p_camp = camp_box.text_frame.paragraphs[0]
    p_camp.text = campaign_name
    p_camp.font.size = Pt(32)
    p_camp.font.bold = True
    p_camp.font.color.rgb = RGBColor(0, 50, 100) # Dark Blue

    # "Proof of Play Pictures" Text
    pop_box = slide.shapes.add_textbox(Inches(0.2), Inches(5.5), Inches(5), Inches(0.5))
    p_pop = pop_box.text_frame.paragraphs[0]
    p_pop.text = "Proof of Play Pictures"
    p_pop.font.size = Pt(28)
    p_pop.font.bold = True
    p_pop.font.color.rgb = RGBColor(0, 50, 100)

    # Note: For the diagonal image collage, you would typically upload the 
    # specific background image to your assets and use:
    # slide.shapes.add_picture('collage_bg.png', Inches(4), 0, width=Inches(6))

# -------------------------
# User Inputs
# -------------------------
campaign_input = st.text_input("📌 Campaign Name")
drive_link = st.text_input("🔗 Google Drive Folder Link")
generate_btn = st.button("🚀 Generate Presentation")

if generate_btn:
    if not campaign_input or not drive_link:
        st.warning("Please fill in all fields")
        st.stop()

    try:
        service = authenticate_drive()
        main_folder_id = extract_folder_id(drive_link)
        subfolders = get_subfolders(service, main_folder_id)

        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)

        # 1. ADD THE FIRST PAGE
        create_title_slide(prs, campaign_input)

        # 2. ADD CONTENT PAGES
        TEAL_COLOR = RGBColor(0, 140, 170) 

        for folder in subfolders:
            folder_name = folder['name']
            images = get_images_in_folder(service, folder['id'])
            
            if not images: continue

            for i in range(0, len(images), 3):
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # Teal Header (Folder Name)
                header_rect = slide.shapes.add_shape(1, Inches(0.2), Inches(0.2), Inches(4.5), Inches(0.7))
                header_rect.fill.solid()
                header_rect.fill.fore_color.rgb = TEAL_COLOR
                header_rect.line.fill.background()
                
                loc_text = header_rect.text_frame
                loc_text.text = folder_name
                loc_text.paragraphs[0].font.size = Pt(18)
                loc_text.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                # Campaign Name Label
                adv_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(6), Inches(0.5))
                adv_box.text_frame.text = f"Campaign Name: {campaign_input}"
                adv_box.text_frame.paragraphs[0].font.size = Pt(22)
                adv_box.text_frame.paragraphs[0].font.bold = True

                # Images
                slide_images = images[i:i+3]
                for idx, img in enumerate(slide_images):
                    img_stream = download_image(service, img["id"])
                    # Standard portrait placement; rotate if required by your previous request
                    pic = slide.shapes.add_picture(img_stream, Inches(0.3 + (idx * 3.2)), Inches(2.2), width=Inches(3.0))

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success("PPT Generated with Title Page!")
        st.download_button(label="📥 Download PPT", data=ppt_io, file_name="Campaign_Report.pptx")

    except Exception as e:
        st.error(f"Error: {e}")
