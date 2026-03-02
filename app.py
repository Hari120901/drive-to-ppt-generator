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

                # Header
                header_rect = slide.shapes.add_shape(
                    1, Inches(0.2), Inches(0.2), Inches(4.5), Inches(0.7)
                )
                header_rect.fill.solid()
                header_rect.fill.fore_color.rgb = TEAL_COLOR
                header_rect.line.fill.background()

                header_text = header_rect.text_frame
                header_text.text = folder_name
                p = header_text.paragraphs[0]
                p.font.bold = True
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.alignment = PP_ALIGN.CENTER

                # Campaign name
                adv_box = slide.shapes.add_textbox(
                    Inches(0.5), Inches(1.1), Inches(6), Inches(0.5)
                )
                adv_para = adv_box.text_frame.paragraphs[0]
                adv_para.text = f"Campaign Name: {campaign_input}"
                adv_para.font.bold = True
                adv_para.font.size = Pt(22)

                # Image Layout
                slide_images = images[i:i + 3]

                column_width = Inches(3.0)
                start_left = Inches(0.4)
                gap = Inches(0.4)
                base_top = Inches(2.2)

                for idx, img in enumerate(slide_images):

                    img_stream = download_image(service, img["id"])
                    left = start_left + (idx * (column_width + gap))

                    picture = slide.shapes.add_picture(
                        img_stream,
                        left,
                        base_top,
                        width=column_width
                    )

                    center_x = picture.left + picture.width // 2
                    center_y = picture.top + picture.height // 2

                    picture.rotation = 90

                    rotated_width = picture.height
                    rotated_height = picture.width

                    picture.left = int(center_x - rotated_width // 2)
                    picture.top = int(center_y - rotated_height // 2)

                    border = slide.shapes.add_shape(
                        1,
                        picture.left,
                        picture.top,
                        rotated_width,
                        rotated_height
                    )
                    border.fill.background()
                    border.line.color.rgb = RGBColor(0, 0, 0)
                    border.line.width = Pt(1.5)

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
