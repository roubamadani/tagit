# Back-end Metadata extraction
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS
import pdfrw
from datetime import datetime
from pdfrw.objects import PdfName, PdfString, PdfDict
import os
import tempfile
from pymediainfo import MediaInfo
import piexif


def get_file_type(uploaded_file):
    """Determine file type based on MIME type"""
    if uploaded_file.type.startswith("image/"):
        return "image"
    elif uploaded_file.type == "application/pdf":
        return "pdf"
    elif uploaded_file.type.startswith("video/"):
        return "video"
    return "unsupported"


def extract_image_metadata(uploaded_file):
    """Extract metadata from image files including EXIF data"""
    img = Image.open(uploaded_file)
    exif_data = img._getexif() or {}

    if exif_data:
        metadata = {
            TAGS.get(tag, tag): value
            for tag, value in exif_data.items()
        }
        if "GPSInfo" in metadata:
            metadata["GPSInfo"] = {
                GPSTAGS.get(key, key): val
                for key, val in metadata["GPSInfo"].items()
            }
    else:
        metadata = {
            "Format": img.format,
            "Mode": img.mode,
            "Size": img.size
        }
    return metadata


def format_metadata(raw_metadata):
    """Format raw metadata into a structured dictionary"""
    if any(raw_metadata.get(key) for key in ["Make", "Model", "Software", "DateTime", "DateTimeOriginal"]):
        formatted = {
            "Basic Info": {
                "Make": raw_metadata.get("Make"),
                "Model": raw_metadata.get("Model"),
                "Software": raw_metadata.get("Software"),
                "DateTime": raw_metadata.get("DateTime"),
                "DateTimeOriginal": raw_metadata.get("DateTimeOriginal")
            },
            "GPS Info": parse_gps(raw_metadata.get("GPSInfo", {})),
        }
    else:
        size = raw_metadata.get("Size")
        if size and isinstance(size, (tuple, list)) and len(size) == 2:
            size_str = f"{size[0]} x {size[1]}"
        else:
            size_str = None

        formatted = {
            "Basic Info": {
                "Format": raw_metadata.get("Format"),
                "Mode": raw_metadata.get("Mode"),
                "Size": size_str
            },
            "GPS Info": None,
        }
    return formatted


def parse_gps(gps):
    """Parse GPS coordinates into human-readable format"""
    if not gps:
        return None
    return {
        "GPSLatitude": f"{gps['GPSLatitude'][0]}° {gps['GPSLatitude'][1]}' {gps['GPSLatitude'][2]}\" {gps['GPSLatitudeRef']}",
        "GPSLongitude": f"{gps['GPSLongitude'][0]}° {gps['GPSLongitude'][1]}' {gps['GPSLongitude'][2]}\" {gps['GPSLongitudeRef']}",
        "GPSTimeStamp": f"{gps['GPSTimeStamp'][0]}:{gps['GPSTimeStamp'][1]}:{gps['GPSTimeStamp'][2]} UTC"
    }


def extract_pdf_metadata(uploaded_file):
    """Extract metadata from PDF files with robust error handling"""
    try:
        uploaded_file.seek(0)
        pdf = pdfrw.PdfReader(uploaded_file)
        raw_info = pdf.Info or {}

        metadata = {}
        for key, value in raw_info.items():
            clean_key = key.lstrip('/') if isinstance(key, str) else str(key)
            metadata[clean_key] = clean_pdf_value(value)

        metadata.update({
            "PDFVersion": str(pdf.Version),
            "PageCount": len(pdf.pages) if hasattr(pdf, 'pages') else 0,
            "ExtractedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        return metadata

    except Exception as e:
        return {"Error": f"Failed to read PDF metadata: {str(e)}"}


def clean_pdf_value(value):
    """Clean and decode PDF metadata values"""
    if value is None:
        return ""

    if isinstance(value, pdfrw.objects.PdfString):
        try:
            return value.decode('utf-8')
        except:
            try:
                return value.decode('latin-1')
            except:
                return str(value)

    if isinstance(value, (bytes, bytearray)):
        try:
            return value.decode('utf-8', errors='replace')
        except:
            return str(value)

    if hasattr(value, '__iter__') and not isinstance(value, str):
        return str([clean_pdf_value(item) for item in value])

    return str(value)


def update_pdf_metadata(uploaded_file, changes):
    """Update PDF metadata and return path to modified file"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            uploaded_file.seek(0)
            tmp_file.write(uploaded_file.read())
            tmp_path = tmp_file.name

        pdf = pdfrw.PdfReader(tmp_path)

        if not pdf.Info:
            pdf.Info = pdfrw.PdfDict()

        for key, value in changes.items():
            if value and str(value).strip():
                pdf.Info[pdfrw.PdfName(key)] = pdfrw.PdfString(str(value))

        output_path = f"modified_{uploaded_file.name}"
        pdfrw.PdfWriter().write(output_path, pdf)

        updated_pdf = pdfrw.PdfReader(output_path)
        if not all(pdfrw.PdfName(k) in updated_pdf.Info for k in changes.keys()):
            raise Exception("Metadata update verification failed")

        os.unlink(tmp_path)

        return output_path

    except Exception as e:
        if 'tmp_path' in locals() and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise Exception(f"Failed to update PDF: {str(e)}")


def extract_video_metadata(uploaded_file):
    """Extract metadata from video files using MediaInfo"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
            uploaded_file.seek(0)
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name

        media_info = MediaInfo.parse(tmp_path)
        os.unlink(tmp_path)

        metadata = {}
        for track in media_info.tracks:
            track_info = {}
            for key, value in track.to_data().items():
                if value:
                    track_info[key] = value
            metadata[track.track_type] = track_info

        return metadata
    except Exception as e:
        return {"Error": f"Failed to extract video metadata: {str(e)}"}


def update_image_metadata(uploaded_file, changes):
    """
    Update image EXIF metadata based on changes dict.
    Supports updates for tags in the 0th IFD (ImageIFD) and ExifIFD.
    Returns path to updated image file.
    """
    try:
        # Create temporary input file
        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1])
        uploaded_file.seek(0)
        tmp_in.write(uploaded_file.read())
        tmp_in.close()

        # Load existing EXIF data
        exif_dict = piexif.load(tmp_in.name)

        # Map common field names to EXIF tags
        tag_map = {
            "Make": piexif.ImageIFD.Make,
            "Model": piexif.ImageIFD.Model,
            "Software": piexif.ImageIFD.Software,
            "DateTime": piexif.ImageIFD.DateTime,
            "DateTimeOriginal": piexif.ExifIFD.DateTimeOriginal,
            "Artist": piexif.ImageIFD.Artist,
            "Copyright": piexif.ImageIFD.Copyright,
            "ImageDescription": piexif.ImageIFD.ImageDescription
        }

        # Apply changes to EXIF data
        for key, value in changes.items():
            if value and key in tag_map:
                if key == "DateTimeOriginal":
                    ifd = "Exif"
                else:
                    ifd = "0th"
                exif_dict[ifd][tag_map[key]] = value.encode('utf-8')

        # Create output file with updated metadata
        tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1])
        piexif.insert(piexif.dump(exif_dict), tmp_in.name, tmp_out.name)
        tmp_out.close()

        # Clean up
        os.unlink(tmp_in.name)

        return tmp_out.name

    except Exception as e:
        # Clean up temporary files if error occurs
        if 'tmp_in' in locals() and os.path.exists(tmp_in.name):
            os.unlink(tmp_in.name)
        if 'tmp_out' in locals() and os.path.exists(tmp_out.name):
            os.unlink(tmp_out.name)
        raise Exception(f"Failed to update image metadata: {str(e)}")