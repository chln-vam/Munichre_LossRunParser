"""
Streamlit Application for Loss Run Parser
- File Upload
- KV Pair Extraction with Confidence Scores
- Editable Cells for Low Confidence Fields
- Export to Duckcreek/Guidewire JSON
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import json
import openpyxl
from io import BytesIO
import base64
from typing import Dict, List, Any

# Import the parser modules
from generalized_loss_run_parser import (
    GeneralizedLossRunParser,
    DuckcreekGuidewireFormatter,
    STANDARD_FIELD_MAPPING
)


# Page Configuration
st.set_page_config(
    page_title="Loss Run Parser",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A5F;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: 600;
        color: #2C5282;
        margin-top: 1rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #C6F6D5;
        border-left: 4px solid #48BB78;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        background-color: #FEFCBF;
        border-left: 4px solid #ECC94B;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        background-color: #FED7D7;
        border-left: 4px solid #F56565;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .metric-card {
        background-color: #EBF8FF;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
    }
    .confidence-high {
        background-color: #C6F6D5;
        color: #22543D;
        padding: 2px 8px;
        border-radius: 4px;
    }
    .confidence-medium {
        background-color: #FEFCBF;
        color: #744210;
        padding: 2px 8px;
        border-radius: 4px;
    }
    .confidence-low {
        background-color: #FED7D7;
        color: #742A2A;
        padding: 2px 8px;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)


class StreamlitLossRunParser:
    """Streamlit UI for Loss Run Parser"""

    def __init__(self):
        self.parser = GeneralizedLossRunParser(output_dir=".")
        self.formatter = DuckcreekGuidewireFormatter()
        self.session_state_init()

    def session_state_init(self):
        """Initialize session state variables"""
        if 'parsed_data' not in st.session_state:
            st.session_state.parsed_data = None
        if 'edited_data' not in st.session_state:
            st.session_state.edited_data = None
        if 'confidence_threshold' not in st.session_state:
            st.session_state.confidence_threshold = 0.80
        if 'file_uploaded' not in st.session_state:
            st.session_state.file_uploaded = False

    def render_header(self):
        """Render application header"""
        st.markdown('<div class="main-header">📊 Loss Run Parser</div>', unsafe_allow_html=True)
        st.markdown("**P&C Insurance Document Processing**")
        st.markdown("---")

    def render_sidebar(self):
        """Render sidebar with configuration"""
        st.sidebar.title("⚙️ Configuration")

        # Confidence threshold slider
        st.sidebar.markdown("### Confidence Threshold")
        st.session_state.confidence_threshold = st.sidebar.slider(
            "Edit cells with confidence below:",
            min_value=0.0,
            max_value=1.0,
            value=0.80,
            step=0.05,
            help="Fields with confidence below this threshold will be editable"
        )

        st.sidebar.markdown("---")

        # Output format selection
        st.sidebar.markdown("### Output Format")
        output_format = st.sidebar.radio(
            "Select format:",
            ["Guidewire", "Rigid Schema"],
            index=0
        )

        st.sidebar.markdown("---")

        # Instructions
        st.sidebar.markdown("""
        ### 📝 Instructions
        1. Upload an Excel file with loss run data
        2. Click "Process File" to extract data
        3. Review extracted KV pairs and confidence scores
        4. Edit fields with low confidence (highlighted in red)
        5. Download corrected JSON output
        """)

        return output_format

    def render_file_upload(self):
        """Render file upload section"""
        st.markdown("### 📁 Upload File")

        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload your loss run Excel file"
        )

        if uploaded_file is not None:
            st.session_state.file_uploaded = True

            # Save uploaded file temporarily
            with open("temp_upload.xlsx", "wb") as f:
                f.write(uploaded_file.getvalue())

            # Show file info
            st.success(f"✅ File uploaded: {uploaded_file.name}")

            # Process button
            if st.button("🚀 Process File", type="primary"):
                self.process_file("temp_upload.xlsx")

        return uploaded_file

    def process_file(self, file_path: str):
        """Process the uploaded file"""
        try:
            with st.spinner("Processing loss run data..."):
                # Parse the file
                parsed_data = self.parser.parse(file_path)

                # Store in session state
                st.session_state.parsed_data = parsed_data
                st.session_state.edited_data = self._deep_copy(parsed_data)

                st.success("✅ File processed successfully!")

        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")

    def _deep_copy(self, data: Dict) -> Dict:
        """Create a deep copy of the parsed data"""
        return json.loads(json.dumps(data))

    def render_summary(self):
        """Render extraction summary"""
        if st.session_state.parsed_data is None:
            return

        data = st.session_state.parsed_data

        st.markdown("### 📊 Extraction Summary")

        # Metadata cards
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold;">{data['lossRunData']['metadata']['total_sheets']}</div>
                <div>Sheets Processed</div>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold;">{data['lossRunData']['metadata']['total_claims']}</div>
                <div>Total Claims</div>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            classification = data['lossRunData']['metadata']['overall_classification']
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 1.5rem; font-weight: bold;">{classification}</div>
                <div>Classification</div>
            </div>
            """, unsafe_allow_html=True)

        with col4:
            confidence = data['lossRunData']['metadata']['classification_confidence']
            confidence_color = "green" if confidence >= 0.95 else "orange" if confidence >= 0.80 else "red"
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold; color: {confidence_color};">{confidence:.1%}</div>
                <div>Confidence</div>
            </div>
            """, unsafe_allow_html=True)

        # Sheet details
        st.markdown("---")
        st.markdown("#### 📋 Sheet Details")

        for sheet in data['lossRunData']['sheets']:
            with st.expander(f"📄 {sheet['sheet_name']}", expanded=True):
                col1, col2 = st.columns(2)

                with col1:
                    st.markdown(f"**Categories Detected:** {', '.join(sheet['classification']['categories_detected'])}")
                    st.markdown(f"**Classification Confidence:** {sheet['classification']['classification_confidence']:.1%}")

                with col2:
                    st.markdown(f"**Total Rows:** {sheet['summary']['total_records']}")
                    st.markdown(f"**Total Columns:** {sheet['classification']['metadata']['total_columns']}")

    def render_extraction_editor(self):
        """Render the KV pair extraction with editable cells"""
        if st.session_state.edited_data is None:
            return

        st.markdown("### 🔍 Extracted Data - Review & Edit")

        threshold = st.session_state.confidence_threshold

        # Process each sheet
        for sheet_idx, sheet in enumerate(st.session_state.edited_data['lossRunData']['sheets']):
            st.markdown(f"#### 📄 Sheet: {sheet['sheet_name']}")

            # Process each claim
            for claim_idx, claim in enumerate(sheet['claims']):
                with st.expander(f"📋 Claim {claim['record_number']} (Row {claim['excel_row']}) - Confidence: {claim['confidence']:.1%}", expanded=True):

                    # Create columns for display
                    cols = st.columns([2, 3, 1, 2])

                    # Header
                    with cols[0]:
                        st.markdown("**Field Name**")
                    with cols[1]:
                        st.markdown("**Value**")
                    with cols[2]:
                        st.markdown("**Confidence**")
                    with cols[3]:
                        st.markdown("**Cell Reference**")

                    st.markdown("---")

                    # Editable fields
                    for field_idx, field in enumerate(claim['fields']):
                        field_name = field['field_name']
                        value = field['value']
                        confidence = field['confidence']
                        cell_ref = field['bounding_box']['cell_reference']
                        original_col = field['original_column']

                        # Determine if editable
                        is_editable = confidence < threshold

                        # Create row
                        col1, col2, col3, col4 = st.columns([2, 3, 1, 2])

                        with col1:
                            # Show original column name for clarity
                            st.markdown(f"**{field_name}**")
                            if field_name != original_col:
                                st.caption(f"({original_col})")

                        with col2:
                            if is_editable:
                                # Editable text input
                                new_value = st.text_input(
                                    f"Value for {field_name}",
                                    value=str(value) if value is not None else "",
                                    key=f"sheet_{sheet_idx}_claim_{claim_idx}_field_{field_idx}",
                                    help=f"Low confidence - edit to correct"
                                )

                                # Update the data
                                try:
                                    # Try to convert to appropriate type
                                    if value is not None:
                                        if isinstance(value, (int, float)):
                                            field['value'] = float(new_value) if '.' in new_value else int(new_value)
                                        else:
                                            field['value'] = new_value if new_value else None
                                    else:
                                        field['value'] = new_value if new_value else None
                                except:
                                    field['value'] = new_value

                                # Mark as manually verified
                                field['confidence'] = 1.0
                                field['manually_corrected'] = True
                            else:
                                # Display non-editable value
                                st.text(str(value) if value is not None else "—")

                        with col3:
                            # Confidence badge
                            if confidence >= 0.95:
                                st.markdown(f'<span class="confidence-high">{confidence:.0%}</span>', unsafe_allow_html=True)
                            elif confidence >= 0.80:
                                st.markdown(f'<span class="confidence-medium">{confidence:.0%}</span>', unsafe_allow_html=True)
                            else:
                                st.markdown(f'<span class="confidence-low">{confidence:.0%}</span>', unsafe_allow_html=True)

                            if is_editable:
                                st.caption("⚠️ Edit me!")

                        with col4:
                            st.code(cell_ref, language=None)

                    # Update claim confidence after edits
                    confidences = [f['confidence'] for f in claim['fields']]
                    claim['confidence'] = sum(confidences) / len(confidences) if confidences else 0.0

        # Save changes button
        if st.button("💾 Save Changes", type="primary"):
            st.success("✅ Changes saved! Download the updated JSON below.")

    def render_export(self, output_format: str):
        """Render export section"""
        if st.session_state.edited_data is None:
            return

        st.markdown("---")
        st.markdown("### 📦 Export")

        # Format output
        if output_format == "Guidewire":
            formatter = DuckcreekGuidewireFormatter()
            output_data = formatter.format(st.session_state.edited_data)
            filename = "loss_run_guidewire_export.json"
        else:
            output_data = st.session_state.edited_data
            filename = "loss_run_rigid_export.json"

        # Convert to JSON
        json_str = json.dumps(output_data, indent=2)

        # Display preview
        with st.expander("👀 JSON Preview", expanded=False):
            st.json(json_str[:2000] + "..." if len(json_str) > 2000 else json_str)

        # Download button
        col1, col2 = st.columns([1, 3])

        with col1:
            st.download_button(
                label=f"⬇️ Download {output_format} JSON",
                data=json_str,
                file_name=filename,
                mime="application/json",
                type="primary"
            )

        with col2:
            # Also generate and show bounding box image
            if st.button("🖼️ Generate Bounding Box Image"):
                with st.spinner("Generating visualization..."):
                    # Re-process to generate image
                    try:
                        # Get the original parsed data for visualization
                        if st.session_state.parsed_data:
                            for sheet in st.session_state.parsed_data['lossRunData']['sheets']:
                                if 'bounding_box_image' in sheet:
                                    st.image(sheet['bounding_box_image'], caption="Bounding Box Visualization")
                    except Exception as e:
                        st.error(f"Error generating image: {str(e)}")

    def render(self):
        """Main render function"""
        self.render_header()

        # Render sidebar and get output format
        output_format = self.render_sidebar()

        # Main content area
        col_main, col_side = st.columns([3, 1])

        with col_main:
            # File upload
            self.render_file_upload()

            # Summary
            if st.session_state.parsed_data:
                self.render_summary()

                # Extraction editor
                self.render_extraction_editor()

                # Export
                self.render_export(output_format)

        # Add some spacing
        st.markdown("<br><br>", unsafe_allow_html=True)


def main():
    """Main application entry point"""
    app = StreamlitLossRunParser()
    app.render()


if __name__ == "__main__":
    main()
