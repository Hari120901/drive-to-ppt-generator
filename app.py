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
st.title("📊 Poster Frames POP PPT Generator")

# -------------------------
# Google Drive Authentication
# -------------------------
def authenticate_drive():
    try:
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gdrive"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        return build("drive", "v3", credentials=credentials)
    except KeyError:
        st.error("Error: 'gdrive' secret not found. Make sure you have configured your secrets in Streamlit.")
        st.stop()


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
        subfolders = get_subfolders(service, main_folder_id)

        if not subfolders:
            st.error("No subfolders found in the provided link.")
            st.stop()

        prs = Presentation()
        # Slide dimensions
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)

        TEAL_COLOR = RGBColor(0, 140, 170)

        for folder in subfolders:
            folder_name = folder["name"]
            folder_id = folder["id"]

            images = get_images_in_folder(service, folder_id)

            if not images:
                continue

            for i in range(0, len(images), 3):

                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # -------------------------
                # Header Rectangle
                # -------------------------
                header_rect = slide.shapes.add_shape(
                    1, Inches(0.2), Inches(0.2), Inches(4.5), Inches(0.7)
                )
                header_rect.fill.solid()
                header_rect.fill.fore_color.rgb = TEAL_COLOR
                header_rect.line.fill.background()

                loc_text = header_rect.text_frame
                loc_text.text = folder_name
                p_loc = loc_text.paragraphs[0]
                p_loc.font.bold = True
                p_loc.font.size = Pt(18)
                p_loc.font.color.rgb = RGBColor(255, 255, 255)
                p_loc.alignment = PP_ALIGN.CENTER

                # -------------------------
                # Campaign Name
                # -------------------------
                adv_box = slide.shapes.add_textbox(
                    Inches(0.5), Inches(1.1), Inches(6), Inches(0.5)
                )
                p_adv = adv_box.text_frame.paragraphs[0]
                p_adv.text = f"Campaign Name: {campaign_input}"
                p_adv.font.bold = True
                p_adv.font.size = Pt(22)
                p_adv.font.color.rgb = RGBColor(0, 0, 0)

                # -------------------------
                # Image Layout (Rotated 90° Clockwise)
                # -------------------------
                slide_images = images[i:i + 3]

                # Adjust visual width to fit 3 images.
                # Since they rotate, they take up their *height* as width on the slide.
                max_visual_width = Inches(3.0)
                # The intended top position for the visual top edge of the rotated image.
                intended_top = Inches(2.2)
                start_left = Inches(0.3)
                gap = Inches(0.3)

                for idx, img in enumerate(slide_images):

                    img_stream = download_image(service, img["id"])
                    
                    # Target center horizontal position
                    center_x = start_left + (idx * (max_visual_width + gap)) + (max_visual_width / 2)

                    # Add image (initially unrotated). Width becomes the visual height.
                    # We don't specify height here to preserve aspect ratio.
                    picture = slide.shapes.add_picture(
                        img_stream,
                        0, # Temporary Left
                        0, # Temporary Top
                        width=max_visual_width
                    )

                    # --- Correcting alignment for 90-degree rotation ---
                    
                    # 1. Rotate 90 degrees clockwise
                    picture.rotation = 90

                    # 2. Get unrotated dimensions (these are what picture.width/height report even after rotation)
                    uw = picture.width # Unrotated Width (becomes visual height)
                    uh = picture.height # Unrotated Height (becomes visual width)

                    # 3. Calculate necessary Left/Top so visual position is correct
                    # Center Y is intended_top + half of visual height
                    center_y = intended_top + (uw / 2)

                    # New Unrotated Left = Center X - (Unrotated Width / 2)
                    new_left = center_x - (uw / 2)
                    # New Unrotated Top = Center Y - (Unrotated Height / 2)
                    new_top = center_y - (uh / 2)

                    picture.left = int(new_left)
                    picture.top = int(new_top)


                    # --- Add matching border ---
                    # The border must be added at unrotated coordinates, then rotated.
                    
                    border = slide.shapes.add_shape(
                        1, # Shape type: Rectangle
                        picture.left,
                        picture.top,
                        picture.width,
                        picture.height
                    )
                    
                    # Configure border appearance
                    border.fill.background() # No fill
                    border.line.color.rgb = RGBColor(0, 0, 0)
                    border.line.width = Pt(1.5)

                    # Important: Rotate border to match picture
                    border.rotation = 90

        # -------------------------
        # Save PPT
        # -------------------------
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
