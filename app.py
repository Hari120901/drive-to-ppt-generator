import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Flipkart Minutes Style PPT Gen", layout="wide")
st.title("📊 Poster Frames POP PPT Generator")

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

                # Header bar
                header_shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(0.1), Inches(0.1), Inches(4.5), Inches(0.7)
                )
                header_shape.fill.solid()
                header_shape.fill.fore_color.rgb = TEAL_COLOR
                header_shape.line.fill.background()
                tf = header_shape.text_frame
                tf.text = folder_name
                p = tf.paragraphs[0]
                p.font.size = Pt(18)
                p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.alignment = PP_ALIGN.CENTER

                # Advertiser text on left top (bold)
                adv_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(7), Inches(0.6))
                adv_para = adv_box.text_frame.paragraphs[0]
                adv_para.text = f"Advertiser: {campaign_input}"
                adv_para.font.size = Pt(24)
                adv_para.font.bold = True
                adv_para.font.color.rgb = RGBColor(0, 0, 0)

                # Flipkart logo on right top
                # If you want to add a logo, you can do so here by loading an image from disk or URL:
                # Example (if logo.png present locally):
                # logo_path = "logo.png"
                # slide.shapes.add_picture(logo_path, prs.slide_width - Inches(1.3), Inches(0.1), width=Inches(1.2))

                slide_images = images[i:i+3]

                img_width = Inches(3.0)
                img_top = Inches(2.0)
                img_left_start = Inches(0.35)
                gap = Inches(0.3)

                for idx, img in enumerate(slide_images):
                    img_stream = download_image(service, img["id"])
                    img_left = img_left_start + idx * (img_width + gap)

                    # Add picture
                    pic = slide.shapes.add_picture(img_stream, img_left, img_top, width=img_width)

                    # Add black border rectangle around the picture
                    border = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        pic.left - Pt(4),  # small negative offset for border thickness
                        pic.top - Pt(4),
                        pic.width + Pt(8),
                        pic.height + Pt(8),
                    )
                    border.fill.background()
                    border.line.color.rgb = RGBColor(0, 0, 0)
                    border.line.width = Pt(2)

                    # Add subtle shadow effect (simulate by adding semi-transparent dark rectangle offset)
                    shadow = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        pic.left + Pt(6),
                        pic.top + Pt(6),
                        pic.width,
                        pic.height,
                    )
                    shadow.fill.solid()
                    shadow.fill.fore_color.rgb = RGBColor(0, 0, 0)
                    shadow.fill.fore_color.alpha = 80  # transparency 0-100 (higher = more transparent)
                    shadow.line.fill.background()
                    # Send shadow behind image and border
                    sp = shadow._element
                    pic._element.addprevious(sp)  # put shadow behind pic
                    border._element.addprevious(sp)  # put shadow behind border

        # Save PPT to bytes
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        st.success("Presentation generated successfully!")

        st.download_button(
            label="📥 Download PPT",
            data=ppt_io,
            file_name=f"{campaign_input}_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        st.error(f"Error: {e}")
