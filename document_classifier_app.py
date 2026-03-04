"""
Intelligent Document Processing Orchestrator
Multi-File Upload with Document Classification Workflow

Workflow:
1. Multi File Upload
2. Document Classifier Model
3. Confidence Score
4. If confidence >= threshold -> Proceed to Parser -> Extraction + Summary
5. If confidence < threshold -> Flag for Review -> User clicks "Ask LLM" -> LLM Semantic Classification -> Correct Type
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import time
import random
import io
from datetime import datetime
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field

# Configure Streamlit page
st.set_page_config(
    page_title="Intelligent Document Processing",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================
# DATA CLASSES
# ============================================

@dataclass
class DocumentMetadata:
    """Store document metadata and processing state"""
    file_obj: Any
    filename: str
    file_type: str
    file_size: int
    detected_type: str = "Unknown"
    confidence: float = 0.0
    status: str = "pending"  # pending, processed, needs_review, llm_corrected
    llm_reasoning: Optional[str] = None
    parsed_data: Optional[Dict] = None
    summary: Optional[str] = None
    upload_time: datetime = field(default_factory=datetime.now)


# ============================================
# MOCK CLASSIFIERS AND PARSERS
# ============================================

class DocumentClassifier:
    """Mock document classifier using keyword-based heuristics"""

    DOCUMENT_TYPES = [
        "Loss Run",
        "Policy Declaration",
        "Invoice",
        "Claim Form",
        "Statement of Value",
       age Account",
        "Unknown"
    ]

    "Broker KEYWORDS = {
        "Loss Run": ["loss", "run", "claims", "incurred", "paid", "reserve", "claimant"],
        "Policy Declaration": ["policy", "declaration", "coverage", "insured", "premium", "limit"],
        "Invoice": ["invoice", "bill to", "payment due", "total due", "billing"],
        "Claim Form": ["claim", "claimant", "incident", "date of loss", "policyholder"],
        "Statement of Value": ["statement of value", "property", "values", "assets", "coverage summary"],
        "Brokerage Account": ["brokerage", "account", "holdings", "securities", "portfolio"]
    }

    def classify(self, text_content: str, filename: str) -> Dict[str, Any]:
        """
        Classify document based on content and filename
        Returns: {detected_type, confidence, reasoning}
        """
        # Simulate processing time
        time.sleep(0.3)

        text_lower = (text_content + " " + filename).lower()

        # Score each document type
        scores = {}
        for doc_type, keywords in self.KEYWORDS.items():
            score = sum(1 for kw in keywords if kw in text_lower)
            # Boost score for filename matches
            if any(kw in filename.lower() for kw in keywords):
                score += 2
            scores[doc_type] = score

        # Find best match
        max_score = max(scores.values()) if scores else 0

        if max_score == 0:
            return {
                "detected_type": "Unknown",
                "confidence": 0.35,
                "reasoning": "No matching keywords found in document"
            }

        # Normalize confidence based on score
        normalized_confidence = min(0.99, 0.5 + (max_score * 0.15))

        # Add some randomness to simulate ML model behavior
        normalized_confidence = normalized_confidence * random.uniform(0.9, 1.0)

        best_type = max(scores, key=scores.get)

        return {
            "detected_type": best_type,
            "confidence": round(normalized_confidence, 3),
            "reasoning": f"Matched keywords: {[kw for kw in self.KEYWORDS[best_type] if kw in text_lower]}"
        }


class LLMClassifier:
    """Mock LLM-powered semantic classifier for ambiguous documents"""

    def __init__(self, api_key: str = None):
        self.api_key = api_key

    def analyze(self, text_content: str, filename: str) -> Dict[str, Any]:
        """
        Use LLM to analyze ambiguous documents
        Returns: {detected_type, confidence, reasoning}
        """
        # Simulate LLM processing time
        time.sleep(1.5)

        # Mock LLM responses based on content analysis
        text_lower = (text_content + " " + filename).lower()

        # Analyze patterns in text
        if "loss" in text_lower or "claim" in text_lower:
            return {
                "detected_type": "Loss Run",
                "confidence": 0.92,
                "reasoning": "Semantic analysis indicates loss run document - found claims-related terminology"
            }
        elif "policy" in text_lower or "coverage" in text_lower:
            return {
                "detected_type": "Policy Declaration",
                "confidence": 0.89,
                "reasoning": "Document contains policy and coverage information"
            }
        elif "invoice" in text_lower or "bill" in text_lower:
            return {
                "detected_type": "Invoice",
                "confidence": 0.94,
                "reasoning": "Billing-related language detected"
            }
        elif "property" in text_lower or "value" in text_lower:
            return {
                "detected_type": "Statement of Value",
                "confidence": 0.87,
                "reasoning": "Property values and asset information present"
            }
        else:
            return {
                "detected_type": "Unknown",
                "confidence": 0.50,
                "reasoning": "Unable to determine document type with high confidence"
            }


class DocumentParser:
    """Mock document parser - extracts data based on document type"""

    def parse(self, file_obj: Any, doc_type: str) -> Dict[str, Any]:
        """
        Parse document and extract relevant data
        """
        time.sleep(0.5)

        # Generate mock extraction results based on document type
        if doc_type == "Loss Run":
            return {
                "carrier": random.choice(["ABC Insurance", "XYZ Mutual", "Global P&C"]),
                "policy_number": f"POL-{random.randint(100000, 999999)}",
                "total_claims": random.randint(5, 50),
                "total_incurred": round(random.uniform(10000, 500000), 2),
                "total_paid": round(random.uniform(5000, 250000), 2),
                "total_reserve": round(random.uniform(5000, 250000), 2),
                "experience_mod": round(random.uniform(0.8, 1.5), 2),
                "losses": [
                    {
                        "claim_number": f"CLM-{i:04d}",
                        "date_of_loss": f"2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}",
                        "description": random.choice(["Liability", "Property", "Auto", "Workers Comp"]),
                        "incurred": round(random.uniform(1000, 50000), 2),
                        "status": random.choice(["Open", "Closed", "Pending"])
                    }
                    for i in range(random.randint(2, 8))
                ]
            }
        elif doc_type == "Policy Declaration":
            return {
                "policy_number": f"POL-{random.randint(100000, 999999)}",
                "named_insured": f"Company {random.choice(['Alpha', 'Beta', 'Gamma', 'Delta'])} Inc.",
                "policy_period": f"2024-01-01 to 2025-01-01",
                "total_premium": round(random.uniform(5000, 100000), 2),
                "coverages": [
                    {
                        "type": random.choice(["General Liability", "Property", "Auto", "Umbrella"]),
                        "limit": f"${random.randint(1,10)}M",
                        "deductible": f"${random.randint(500, 25000)}"
                    }
                    for _ in range(random.randint(2, 5))
                ]
            }
        elif doc_type == "Invoice":
            return {
                "invoice_number": f"INV-{random.randint(1000, 9999)}",
                "invoice_date": f"2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}",
                "due_date": f"2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}",
                "bill_to": f"Client {random.choice(['A', 'B', 'C', 'D'])} Corp",
                "total_amount": round(random.uniform(1000, 50000), 2),
                "line_items": [
                    {
                        "description": random.choice(["Service Fee", "Premium Payment", "Consulting"]),
                        "quantity": random.randint(1, 10),
                        "unit_price": round(random.uniform(100, 1000), 2)
                    }
                    for _ in range(random.randint(2, 5))
                ]
            }
        else:
            return {
                "status": "parsed",
                "document_type": doc_type,
                "raw_text_preview": "Sample extracted text from document...",
                "extraction_confidence": round(random.uniform(0.7, 0.95), 2)
            }

    def generate_summary(self, parsed_data: Dict, doc_type: str) -> str:
        """Generate human-readable summary of parsed data"""
        if doc_type == "Loss Run":
            total_claims = parsed_data.get("total_claims", 0)
            total_incurred = parsed_data.get("total_incurred", 0)
            return f"Loss Run document with {total_claims} claims. Total incurred: ${total_incurred:,.2f}."
        elif doc_type == "Policy Declaration":
            insured = parsed_data.get("named_insured", "Unknown")
            premium = parsed_data.get("total_premium", 0)
            return f"Policy Declaration for {insured}. Total Premium: ${premium:,.2f}."
        elif doc_type == "Invoice":
            invoice_num = parsed_data.get("invoice_number", "N/A")
            amount = parsed_data.get("total_amount", 0)
            return f"Invoice #{invoice_num}. Total Amount: ${amount:,.2f}."
        else:
            return f"Document type: {doc_type}. Extraction completed."


# ============================================
# STREAMLIT UI COMPONENTS
# ============================================

class IDPOrchestratorApp:
    """Main Streamlit Application"""

    def __init__(self):
        self.classifier = DocumentClassifier()
        self.llm_classifier = LLMClassifier()
        self.parser = DocumentParser()
        self.init_session_state()

    def init_session_state(self):
        """Initialize session state variables"""
        if 'documents' not in st.session_state:
            st.session_state.documents = {}
        if 'confidence_threshold' not in st.session_state:
            st.session_state.confidence_threshold = 0.75
        if 'processing_complete' not in st.session_state:
            st.session_state.processing_complete = False
        if 'extraction_results' not in st.session_state:
            st.session_state.extraction_results = {}

    def render_custom_css(self):
        """Render custom CSS styling"""
        st.markdown("""
        <style>
            .main-header {
                font-size: 2.5rem;
                font-weight: bold;
                color: #1E3A5F;
                margin-bottom: 0.5rem;
            }
            .sub-header {
                font-size: 1.5rem;
                font-weight: 600;
                color: #2C5282;
                margin-top: 1.5rem;
                margin-bottom: 1rem;
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
            .info-box {
                padding: 1rem;
                background-color: #BEE3F8;
                border-left: 4px solid #4299E1;
                border-radius: 4px;
                margin: 1rem 0;
            }
            .metric-card {
                background-color: #EBF8FF;
                padding: 1rem;
                border-radius: 8px;
                text-align: center;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .status-badge {
                padding: 0.25rem 0.75rem;
                border-radius: 9999px;
                font-size: 0.875rem;
                font-weight: 500;
            }
            .status-processed {
                background-color: #C6F6D5;
                color: #22543D;
            }
            .status-review {
                background-color: #FEFCBF;
                color: #744210;
            }
            .status-llm {
                background-color: #BEE3F8;
                color: #2A4365;
            }
            .status-pending {
                background-color: #E2E8F0;
                color: #4A5568;
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
            .doc-type-highlight {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                font-weight: bold;
            }
            div.stButton > button {
                width: 100%;
            }
            .workflow-step {
                display: flex;
                align-items: center;
                padding: 0.5rem;
                margin: 0.25rem 0;
                border-radius: 4px;
                background-color: #F7FAFC;
            }
            .workflow-arrow {
                margin: 0 0.5rem;
                color: #A0AEC0;
            }
        </style>
        """, unsafe_allow_html=True)

    def render_header(self):
        """Render application header"""
        st.markdown('<div class="main-header">📄 Intelligent Document Processing</div>', unsafe_allow_html=True)
        st.markdown("**Multi-File Upload → Classification → Confidence Scoring → Parsing → Extraction**")
        st.markdown("---")

    def render_workflow_diagram(self):
        """Render workflow diagram in sidebar"""
        st.sidebar.markdown("""
        ### 🔄 Workflow Diagram

        ```mermaid
        graph LR
            A[📁 Multi File Upload] --> B[🔍 Document Classifier]
            B --> C{📊 Confidence Score}
            C -->|≥ Threshold| D[✅ Proceed to Parser]
            C -->|< Threshold| E[🚩 Flag for Review]
            D --> F[📝 Extraction + Summary]
            E --> G[🧠 Ask LLM]
            G --> H[📋 LLM Semantic Classification]
            H --> I[✓ Correct Type]
            I --> D
        ```
        """, unsafe_allow_html=True)

    def render_sidebar(self):
        """Render sidebar with configuration"""
        self.render_workflow_diagram()

        st.sidebar.title("⚙️ Configuration")

        # Confidence threshold slider
        st.sidebar.markdown("### Confidence Threshold")
        st.session_state.confidence_threshold = st.sidebar.slider(
            "Minimum confidence to auto-process:",
            min_value=0.0,
            max_value=1.0,
            value=0.75,
            step=0.05,
            help="Documents with confidence below this threshold will be flagged for review"
        )

        st.sidebar.markdown("---")

        # LLM Settings
        st.sidebar.markdown("### LLM Settings")
        llm_api_key = st.sidebar.text_input(
            "OpenAI API Key",
            type="password",
            help="Required for LLM semantic classification"
        )

        if llm_api_key:
            st.sidebar.success("✅ API Key configured")
        else:
            st.sidebar.info("ℹ️ LLM will use mock responses")

        st.sidebar.markdown("---")

        # Statistics
        st.sidebar.markdown("### 📊 Session Statistics")

        if st.session_state.documents:
            total = len(st.session_state.documents)
            processed = sum(1 for d in st.session_state.documents.values() if d.status in ["processed", "llm_corrected"])
            review = sum(1 for d in st.session_state.documents.values() if d.status == "needs_review")

            st.sidebar.metric("Total Files", total)
            st.sidebar.metric("Auto-Processed", processed, delta=f"{processed/total*100:.0f}%" if total > 0 else "0%")
            st.sidebar.metric("Needs Review", review, delta=f"-{review}" if review > 0 else "0")
        else:
            st.sidebar.info("No files uploaded yet")

        st.sidebar.markdown("---")

        # Reset button
        if st.sidebar.button("🔄 Reset Session", type="secondary"):
            st.session_state.documents = {}
            st.session_state.extraction_results = {}
            st.session_state.processing_complete = False
            st.rerun()

        # Instructions
        st.sidebar.markdown("""
        ### 📝 Instructions

        1. **Upload** multiple documents (PDF, Excel, CSV, Images)
        2. **Classification** runs automatically on each file
        3. **High Confidence** (≥ threshold) → Auto-parsed
        4. **Low Confidence** (< threshold) → Flagged for review
        5. **Review** flagged files → Click "Ask LLM" for semantic classification
        6. **Export** final extraction results
        """)

        return llm_api_key

    def render_file_upload(self):
        """Render file upload section"""
        st.markdown("### 📁 Multi-File Upload")

        uploaded_files = st.file_uploader(
            "Drop your documents here",
            type=['pdf', 'xlsx', 'xls', 'csv', 'png', 'jpg', 'jpeg'],
            accept_multiple_files=True,
            help="Supported formats: PDF, Excel, CSV, Images"
        )

        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")

            # Process new files
            new_files = [f for f in uploaded_files if f.name not in st.session_state.documents]

            if new_files:
                with st.spinner("Classifying documents..."):
                    for file in new_files:
                        self.process_document(file)

                st.rerun()

        return uploaded_files

    def extract_file_content(self, file_obj) -> str:
        """Extract text content from uploaded file for classification"""
        try:
            # Read file content
            content = file_obj.getvalue()

            # For demo purposes, we'll use filename as a proxy for content
            # In production, you would use OCR (pytesseract) or PDF text extraction
            filename = file_obj.name

            # Create mock text content based on filename
            text_content = filename + " "

            # Add some mock content based on file extension
            if file_obj.name.endswith('.xlsx') or file_obj.name.endswith('.xls'):
                text_content += "spreadsheet data rows columns Excel workbook sheet"
            elif file_obj.name.endswith('.pdf'):
                text_content += "document pdf portable format text"
            elif file_obj.name.endswith('.csv'):
                text_content += "comma separated values data table"
            elif file_obj.name.endswith(('.png', '.jpg', '.jpeg')):
                text_content += "image scanned document picture"

            return text_content

        except Exception as e:
            st.error(f"Error extracting content: {str(e)}")
            return file_obj.name

    def process_document(self, file_obj):
        """Process a single document through classification"""
        # Extract content for classification
        text_content = self.extract_file_content(file_obj)

        # Classify document
        classification_result = self.classifier.classify(text_content, file_obj.name)

        # Determine status based on confidence threshold
        threshold = st.session_state.confidence_threshold
        if classification_result['confidence'] >= threshold:
            status = "processed"
        else:
            status = "needs_review"

        # Create document metadata
        doc_meta = DocumentMetadata(
            file_obj=file_obj,
            filename=file_obj.name,
            file_type=file_obj.type or "unknown",
            file_size=file_obj.size,
            detected_type=classification_result['detected_type'],
            confidence=classification_result['confidence'],
            status=status,
            llm_reasoning=classification_result['reasoning']
        )

        # Store in session state
        st.session_state.documents[file_obj.name] = doc_meta

    def render_dashboard(self):
        """Render the processing dashboard"""
        if not st.session_state.documents:
            return

        st.markdown("### 📊 Processing Dashboard")

        # Summary metrics
        total = len(st.session_state.documents)
        processed = sum(1 for d in st.session_state.documents.values() if d.status == "processed")
        review = sum(1 for d in st.session_state.documents.values() if d.status == "needs_review")
        llm_corrected = sum(1 for d in st.session_state.documents.values() if d.status == "llm_corrected")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold;">{total}</div>
                <div>Total Files</div>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold; color: #48BB78;">{processed}</div>
                <div>Auto-Processed</div>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold; color: #ECC94B;">{review}</div>
                <div>Needs Review</div>
            </div>
            """, unsafe_allow_html=True)

        with col4:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 2rem; font-weight: bold; color: #4299E1;">{llm_corrected}</div>
                <div>LLM Corrected</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")

        # Document table
        self.render_document_table()

    def render_document_table(self):
        """Render interactive document table"""
        st.markdown("#### 📋 Document Status")

        # Prepare data for table
        table_data = []
        for doc in st.session_state.documents.values():
            # Re-evaluate status based on current threshold
            if doc.status not in ["llm_corrected"]:
                if doc.confidence >= st.session_state.confidence_threshold:
                    doc.status = "processed"
                else:
                    doc.status = "needs_review"

            table_data.append({
                "Filename": doc.filename,
                "Type": doc.detected_type,
                "Confidence": doc.confidence,
                "Status": doc.status,
                "Size (KB)": f"{doc.file_size / 1024:.1f}"
            })

        df = pd.DataFrame(table_data)

        # Display with styling
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

    def render_review_section(self, llm_api_key: str):
        """Render the review section for flagged documents"""
        flagged_docs = [d for d in st.session_state.documents.values() if d.status == "needs_review"]

        if not flagged_docs:
            return

        st.markdown("---")
        st.markdown("### 🚩 Documents Needing Review")

        st.warning(f"⚠️ **{len(flagged_docs)} document(s)** flagged for review due to low confidence scores")

        for doc in flagged_docs:
            with st.expander(f"📄 {doc.filename} - {doc.detected_type} ({doc.confidence:.1%})"):
                col1, col2 = st.columns([2, 1])

                with col1:
                    st.markdown(f"**Detected Type:** {doc.detected_type}")
                    st.markdown(f"**Confidence:** {doc.confidence:.1%}")
                    st.markdown(f"**Reasoning:** {doc.llm_reasoning or 'N/A'}")

                    # Manual override dropdown
                    new_type = st.selectbox(
                        "Override document type:",
                        options=["Loss Run", "Policy Declaration", "Invoice", "Claim Form",
                                "Statement of Value", "Brokerage Account", "Unknown"],
                        index=["Loss Run", "Policy Declaration", "Invoice", "Claim Form",
                               "Statement of Value", "Brokerage Account", "Unknown"].index(doc.detected_type)
                               if doc.detected_type in ["Loss Run", "Policy Declaration", "Invoice", "Claim Form",
                               "Statement of Value", "Brokerage Account", "Unknown"] else 6,
                        key=f"type_override_{doc.filename}"
                    )

                    if new_type != doc.detected_type:
                        if st.button(f"Apply '{new_type}' to {doc.filename}", key=f"apply_{doc.filename}"):
                            doc.detected_type = new_type
                            doc.status = "llm_corrected"
                            doc.llm_reasoning = f"Manually overridden to {new_type}"
                            st.success(f"Updated {doc.filename} to {new_type}")
                            st.rerun()

                with col2:
                    # Ask LLM button
                    st.markdown("**Need semantic analysis?**")

                    if st.button(f"🧠 Ask LLM to Analyze", key=f"llm_{doc.filename}"):
                        with st.spinner("LLM is analyzing document..."):
                            # Extract content for LLM
                            text_content = self.extract_file_content(doc.file_obj)

                            # Run LLM classification
                            llm_result = self.llm_classifier.analyze(text_content, doc.filename)

                            # Update document
                            doc.detected_type = llm_result['detected_type']
                            doc.confidence = llm_result['confidence']
                            doc.llm_reasoning = llm_result['reasoning']
                            doc.status = "llm_corrected"

                            st.success(f"LLM classified as: {llm_result['detected_type']}")
                            st.rerun()

    def render_extraction_section(self):
        """Render extraction and parsing section"""
        # Get documents ready for parsing
        ready_docs = [d for d in st.session_state.documents.values()
                     if d.status in ["processed", "llm_corrected"]]

        if not ready_docs:
            return

        st.markdown("---")
        st.markdown("### 📝 Extraction & Summary")

        # Parse button
        col1, col2 = st.columns([1, 3])

        with col1:
            if st.button(f"🚀 Parse {len(ready_docs)} Document(s)", type="primary"):
                with st.spinner("Extracting data from documents..."):
                    for doc in ready_docs:
                        # Parse document
                        parsed_data = self.parser.parse(doc.file_obj, doc.detected_type)
                        doc.parsed_data = parsed_data

                        # Generate summary
                        doc.summary = self.parser.generate_summary(parsed_data, doc.detected_type)

                        # Store in extraction results
                        st.session_state.extraction_results[doc.filename] = {
                            "document_type": doc.detected_type,
                            "parsed_data": parsed_data,
                            "summary": doc.summary,
                            "confidence": doc.confidence
                        }

                    st.session_state.processing_complete = True
                    st.success(f"✅ Successfully parsed {len(ready_docs)} document(s)")
                    st.rerun()

        # Display results
        if st.session_state.extraction_results:
            self.render_parsed_results()

    def render_parsed_results(self):
        """Render parsed extraction results"""
        st.markdown("#### 📊 Parsed Results")

        # Create tabs for different views
        tab1, tab2, tab3 = st.tabs(["📋 Summary View", "🔍 Detailed View", "📦 Export"])

        with tab1:
            # Summary cards
            for filename, result in st.session_state.extraction_results.items():
                with st.expander(f"📄 {filename} - {result['document_type']}"):
                    st.markdown(f"**Summary:** {result['summary']}")

                    confidence = result['confidence']
                    if confidence >= 0.9:
                        st.markdown(f"**Confidence:** <span class='confidence-high'>{confidence:.1%}</span>",
                                   unsafe_allow_html=True)
                    elif confidence >= 0.75:
                        st.markdown(f"**Confidence:** <span class='confidence-medium'>{confidence:.1%}</span>",
                                   unsafe_allow_html=True)
                    else:
                        st.markdown(f"**Confidence:** <span class='confidence-low'>{confidence:.1%}</span>",
                                   unsafe_allow_html=True)

        with tab2:
            # Detailed JSON view
            for filename, result in st.session_state.extraction_results.items():
                with st.expander(f"📄 {filename}"):
                    st.json(result['parsed_data'])

        with tab3:
            # Export options
            self.render_export_options()

    def render_export_options(self):
        """Render export options"""
        st.markdown("#### 📦 Export Options")

        # Convert results to DataFrame for export
        export_data = []
        for filename, result in st.session_state.extraction_results.items():
            row = {
                "Filename": filename,
                "Document Type": result['document_type'],
                "Summary": result['summary'],
                "Confidence": f"{result['confidence']:.1%}"
            }

            # Flatten parsed data
            if result['parsed_data']:
                for key, value in result['parsed_data'].items():
                    if not isinstance(value, (list, dict)):
                        row[key] = str(value)

            export_data.append(row)

        df = pd.DataFrame(export_data)

        # Display preview
        st.dataframe(df, use_container_width=True)

        # Download buttons
        col1, col2 = st.columns(2)

        with col1:
            # CSV download
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name="extraction_results.csv",
                mime="text/csv"
            )

        with col2:
            # JSON download
            json_data = json.dumps(st.session_state.extraction_results, indent=2, default=str)
            st.download_button(
                label="📥 Download as JSON",
                data=json_data,
                file_name="extraction_results.json",
                mime="application/json"
            )

    def render(self):
        """Main render function"""
        self.render_custom_css()
        self.render_header()

        # Sidebar
        llm_api_key = self.render_sidebar()

        # Main content
        uploaded_files = self.render_file_upload()

        if st.session_state.documents:
            # Dashboard
            self.render_dashboard()

            # Review section
            self.render_review_section(llm_api_key)

            # Extraction section
            self.render_extraction_section()


# ============================================
# MAIN APPLICATION
# ============================================

def main():
    """Main application entry point"""
    app = IDPOrchestratorApp()
    app.render()


if __name__ == "__main__":
    main()
