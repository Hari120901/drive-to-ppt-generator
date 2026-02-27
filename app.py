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
st.title("📊 Professional PPT Generator (Folder-wise Grouping with Rotated Images)")

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
    """Fetch all subfolders within the provided link"""
    query = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    return results.get("files", [])

def get_images_in_folder(service, folder_id):
    """Fetch images only from a specific folder"""
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
# User Inputs
# -------------------------
campaign_input = st.text_input("📌 Campaign Name")
drive_link = st.text_input("🔗 Google Drive Folder Link (containing subfolders)")
generate_btn = st.button("🚀 Generate Presentation")

if generate_btn:
    if not campaign_input or not drive_link:
        st.warning("Please fill in all fields")
        st.stop()

    try:
        service = authenticate_drive()
        main_folder_id = extract_folder_id(drive_link)
        
        # Get all subfolders (e.g., Omaxe The Palace, etc.)
        subfolders = get_subfolders(service, main_folder_id)
        
        if not subfolders:
            st.error("No subfolders found in the provided link.")
            st.stop()

        prs = Presentation()
        # Set slide dimensions to Standard (10 x 7.5) to match the screenshot aspect ratio
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)

        # UI Constants
        TEAL_COLOR = RGBColor(0, 140, 170) 

        for folder in subfolders:
            folder_name = folder['name']
            folder_id = folder['id']
            
            # Fetch images for THIS specific folder
            images = get_images_in_folder(service, folder_id)
            
            if not images:
                continue # Skip folders with no images

            # Split images into groups of 3 (for the 3-column layout)
            for i in range(0, len(images), 3):
                slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank Slide

                # 1. Teal Header Rectangle (Folder Name)
                header_rect = slide.shapes.add_shape(
                    1, # Rectangle
                    Inches(0.2), Inches(0.2), Inches(4.5), Inches(0.7)
                )
                header_rect.fill.solid()
                header_rect.fill.fore_color.rgb = TEAL_COLOR
                header_rect.line.fill.background() # No border

                loc_text = header_rect.text_frame
                loc_text.text = folder_name
                p_loc = loc_text.paragraphs[0]
                p_loc.font.bold = True
                p_loc.font.size = Pt(18)
                p_loc.font.color.rgb = RGBColor(255, 255, 255)
                p_loc.alignment = PP_ALIGN.CENTER

                # 2. Campaign Name Label
                adv_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(6), Inches(0.5))
                p_adv = adv_box.text_frame.paragraphs[0]
                p_adv.text = f"Campaign Name: {campaign_input}"
                p_adv.font.bold = True
                p_adv.font.size = Pt(22)
                p_adv.font.color.rgb = RGBColor(0, 0, 0)

                # 3. Add 3 Images horizontally (Rotated 90 degrees Clockwise)
                slide_images = images[i:i+3]
                
                # New rotated dimensions (swapped and scaled down slightly to fit 3 in a row)
                img_width = Inches(3.0) # Original width of portrait picture
                img_height = Inches(4.5) # Original height of portrait picture

                # Layout definitions
                start_left = Inches(0.3)
                gap = Inches(0.2)
                top_pos = Inches(2.2)

                for idx, img in enumerate(slide_images):
                    img_stream = download_image(service, img["id"])
                    left = start_left + (idx * (img_width + gap))
                    
                    # Add Picture with 90 degree clockwise rotation (horizontal orientation)
                    # We specify the standard height/width and then rotate the container
                    # pic = slide.shapes.add_picture(img_stream, left, top_pos, width=img_width, height=img_height)
                    # pic.rotation = 90
                    
                    # Since we want them rotated but still within a black frame, we will rotate the picture and frame together using a grouped shape, but pptx doesn't handle rotation of grouped pictures well.
                    
                    # We can directly set the dimensions to portrait style, and rotation to horizontal style.
                    pic = slide.shapes.add_picture(img_stream, left, top_pos, width=img_width, height=img_height)
                    pic.rotation = -90 # Negative 90 for landscape orientation (top edge to the left)
                    
                    # Add Black Border/Frame around the portrait oriented picture (rotated shape)
                    rect = slide.shapes.add_shape(1, left, top_pos, img_width, img_height)
                    rect.fill.background()
                    rect.line.color.rgb = RGBColor(0, 0, 0)
                    rect.line.width = Pt(1.5)
                    rect.rotation = -90 # Rotate border too to match image shape

        # Save and Download
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success("Presentation generated successfully!")
        st.download_button(
            label="📥 Download PPT",
            data=ppt_io,
            file_name=f"{campaign_input}_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        st.error(f"Error: {e}")
