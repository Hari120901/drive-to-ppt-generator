import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Poster Frames PPT", layout="wide")
st.title("📊 Poster Frames POP PPT Generator")

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
        st.error("Invalid Google Drive folder link")
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
drive_link = st.text_input("🔗 Google Drive Folder Link")
generate_btn = st.button("🚀 Generate Presentation")


if generate_btn:

    if not campaign_input or not drive_link:
        st.warning("Please fill all fields")
        st.stop()

    try:

        service = authenticate_drive()
        main_folder_id = extract_folder_id(drive_link)
        subfolders = get_subfolders(service, main_folder_id)

        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        # -------------------------
        # Colors
        # -------------------------
        TEAL = RGBColor(0, 150, 160)
        GREY = RGBColor(242, 242, 242)
        FRENCH_NAVY = RGBColor(11, 35, 65)

        # -------------------------
        # Image Layout
        # -------------------------
        image_width = Inches(3)
        gap = Inches(1.5)
        top_position = Inches(2.2)

        left_positions = [
            Inches(1.3),
            Inches(1.3) + image_width + gap
        ]

        # -------------------------
        # Loop Through Subfolders
        # -------------------------
        for folder in subfolders:

            images = get_images_in_folder(service, folder["id"])

            if not images:
                continue

            for i in range(0, len(images), 2):

                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # -------------------------
                # Background
                # -------------------------
                bg = slide.background
                fill = bg.fill
                fill.solid()
                fill.fore_color.rgb = GREY

                # -------------------------
                # Campaign Label
                # -------------------------
                label_box = slide.shapes.add_textbox(
                    Inches(0.9),
                    Inches(0.6),
                    Inches(5),
                    Inches(0.6)
                )

                tf = label_box.text_frame
                p = tf.paragraphs[0]
                p.text = "C A M P A I G N  N A M E:"
                p.font.size = Pt(26)
                p.font.name = "Montserrat"
                p.font.bold = False
                p.font.color.rgb = FRENCH_NAVY

                # -------------------------
                # Campaign Name
                # -------------------------
                name_box = slide.shapes.add_textbox(
                    Inches(0.9),
                    Inches(1.0),
                    Inches(7),
                    Inches(1)
                )

                tf = name_box.text_frame
                p = tf.paragraphs[0]
                p.text = campaign_input
                p.font.size = Pt(40)
                p.font.name = "Montserrat"
                p.font.bold = True
                p.font.color.rgb = FRENCH_NAVY

                # -------------------------
                # Store / Folder Name
                # -------------------------
                store_box = slide.shapes.add_textbox(
                    Inches(8.5),
                    Inches(0.8),
                    Inches(4),
                    Inches(1)
                )

                tf = store_box.text_frame
                p = tf.paragraphs[0]
                p.text = folder["name"].upper()
                p.font.size = Pt(36)
                p.font.name = "Montserrat"
                p.font.bold = True
                p.font.color.rgb = TEAL
                p.alignment = PP_ALIGN.RIGHT

                # -------------------------
                # Add Images
                # -------------------------
                slide_images = images[i:i+2]

                for idx, img in enumerate(slide_images):

                    img_stream = download_image(service, img["id"])

                    picture = slide.shapes.add_picture(
                        img_stream,
                        left_positions[idx],
                        top_position,
                        width=image_width
                    )

                    # Frame Border
                    border = slide.shapes.add_shape(
                        1,
                        picture.left,
                        picture.top,
                        picture.width,
                        picture.height
                    )

                    border.fill.background()
                    border.line.color.rgb = TEAL
                    border.line.width = Pt(3)

        # -------------------------
        # Save PPT
        # -------------------------
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success("✅ Presentation Generated Successfully!")

        st.download_button(
            label="📥 Download PPT",
            data=ppt_io,
            file_name=f"{campaign_input}_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        st.error(f"Error: {e}")
