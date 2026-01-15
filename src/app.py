"""
Streamlit Application - Web Interface for Vertify Mapping Generator.

This module contains only the user interface logic.
Business logic is separated into specific modules.
"""

import json
import sys
from pathlib import Path
import streamlit as st

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent))

from generator import MappingSpreadsheetGenerator


def configure_page():
    """Configures Streamlit page properties."""
    st.set_page_config(
        page_title="Vertify Mapping Generator",
        page_icon="📊",
        layout="wide"
    )


def render_header():
    """Renders the application header."""
    st.title("📊 Vertify Mapping Spreadsheet Generator")
    st.markdown("**Convert Vertify mapping JSON files into formatted Excel spreadsheets**")
    st.divider()


def render_file_uploader():
    """
    Renders the file upload component.
    
    Returns:
        UploadedFile or None: Uploaded file or None
    """
    return st.file_uploader(
        "📁 Upload mapping JSON file",
        type=["json"],
        help="Select the JSON file exported from Vertify"
    )


def render_statistics(stats):
    """
    Renders the processed JSON statistics.
    
    Args:
        stats: Dictionary with statistics
    """
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("📋 ObjectMaps", stats["total_objectmaps"])
    
    with col2:
        st.metric("🔄 Total Properties", stats["total_properties"])
    
    with col3:
        st.metric("🔍 Total Filters", stats["total_filters"])


def render_preview_table(objects_map):
    """
    Renders the ObjectMaps preview table.
    
    Args:
        objects_map: List of ObjectMaps from JSON
    """
    with st.expander("👀 ObjectMaps Preview", expanded=True):
        if objects_map:
            preview_data = []
            for idx, obj in enumerate(objects_map, 1):
                preview_data.append({
                    "ID": idx,
                    "Name": obj.get("Name", "N/A"),
                    "Source": obj.get("SourceSystemName", "N/A"),
                    "Target": obj.get("TargetSystemName", "N/A"),
                    "Properties": len(obj.get("PropertiesMap", [])),
                    "Filters": len(obj.get("ObjectsMapFilter", [])),
                })
            
            st.dataframe(preview_data, use_container_width=True)
        else:
            st.warning("No ObjectMap found in JSON")


def generate_and_download(json_data, uploaded_file):
    """
    Generates the spreadsheet and provides automatic download.
    
    Args:
        json_data: Loaded JSON data
        uploaded_file: Uploaded file
    """
    with st.spinner("Generating spreadsheet... Please wait..."):
        try:
            # Generate spreadsheet
            generator = MappingSpreadsheetGenerator(json_data)
            excel_bytes = generator.generate_to_bytes()
            
            # Output filename
            output_filename = uploaded_file.name.replace('.json', '_MAPPINGS.xlsx')
            
            st.success("✅ Spreadsheet generated successfully!")
            
            # Download button
            st.download_button(
                label="⬇️ Download Excel Spreadsheet",
                data=excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch'
            )
            
        except Exception as e:
            st.error(f"❌ Error generating spreadsheet: {str(e)}")
            with st.expander("Error details"):
                st.exception(e)


def render_instructions():
    """Renders usage instructions when no file is uploaded."""
    st.info("👆 Upload a JSON file to begin")
    
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        ### Step by step:
        
        1. **Export** the Vertify mapping JSON file
        2. **Upload** the file using the field above
        3. **Review** the displayed information
        4. The spreadsheet will be **generated automatically**
        5. **Download** the generated XLSX file
        
        ### What the spreadsheet contains:
        
        - **Tab 1**: Summary of all movements
        - **Tabs 2-N**: Details of each ObjectMap including:
          - API Request configuration
          - Merge rules
          - Filter conditions
          - Field mappings
        
        ### ⚠️ Limitations - Fields to be completed manually:
        
        The following fields are **not available** in the original JSON and will be generated with placeholder values:
        
        **Tab 1 - Movements to migrate:**
        - **Trigger Type** → Default: "Collect & Move"
        - **Interval frequency** → Default: "at 00:00 AM"
        - **Interval days** → Default: "every ?"
        - **Sandbox** (Source/Target) → Default: "TRUE/FALSE"
        - **Credentials** (Source/Target) → Default: "TRUE/FALSE"
        - **Customization** → Default: "TRUE/FALSE"
        - **Notes** → Empty
        - **Email Alert** → Empty
        - **Email Every** → Empty
        
        **Detail Tabs - API Request:**
        - **Type** → Default: "REST"
        - **Path/connection string** → Empty
        - **Request example** → Empty
        - **Response example** → Empty
        - **Notes** → Empty
        
        **Note:** These fields must be filled in manually in the generated Excel spreadsheet based on your specific requirements.
        """)


def render_footer():
    """Renders the application footer."""
    st.divider()
    st.markdown(
        "<div style='text-align: center; color: gray;'>Digibee</div>",
        unsafe_allow_html=True
    )


def process_uploaded_file(uploaded_file):
    """
    Processes the uploaded JSON file and renders appropriate content.
    
    Args:
        uploaded_file: File uploaded by the user
    """
    try:
        # Read and parse JSON
        json_data = json.load(uploaded_file)
        
        # Create generator to get statistics
        generator = MappingSpreadsheetGenerator(json_data)
        stats = generator.get_statistics()
        
        # Render statistics
        render_statistics(stats)
        
        st.divider()
        
        # ObjectMaps preview
        objects_map = json_data.get("ObjectsMap", [])
        render_preview_table(objects_map)
        
        st.divider()
        
        # Auto-generate and provide download
        generate_and_download(json_data, uploaded_file)
    
    except json.JSONDecodeError as e:
        st.error("❌ Error reading JSON: Invalid file")
        with st.expander("Error details"):
            st.exception(e)
    
    except Exception as e:
        st.error(f"❌ Unexpected error: {str(e)}")
        with st.expander("Error details"):
            st.exception(e)


def main():
    """Main application function."""
    configure_page()
    render_header()
    
    # File upload
    uploaded_file = render_file_uploader()
    
    if uploaded_file is not None:
        process_uploaded_file(uploaded_file)
    else:
        render_instructions()
    
    render_footer()


if __name__ == "__main__":
    main()
