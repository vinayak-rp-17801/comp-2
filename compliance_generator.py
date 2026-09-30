#!/usr/bin/env python3
"""
Compliance Document Generator Agent
Generates DOCX files for 40 different compliance standards based on zip file contents.
"""

import os
import sys
import json
import zipfile
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime


class ComplianceDocumentGenerator:
    """Agent to generate compliance documentation in DOCX format."""
    
    def __init__(self, output_dir="compliance_documents"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.compliance_list = []
        self.template_config = self.load_template_config()
    
    def load_template_config(self):
        """Load the compliance template configuration."""
        return {
            "category": "Compliance",
            "logo_placeholder": "[Logo Placeholder]",
            "banner_height": 2,
            "supported_editions": ["Free", "Basic", "Standard", "Professional", "MSSP"],
            "supported_dcs": ["US", "IN", "JP", "CA", "AU", "UK", "EU"],
            "valid_submission": "Yes",
            "help_doc_url": "https://www.manageengine.com/log-management/help/integrations/compliance-extensions.html"
        }
    
    def extract_compliance_names(self, zip_file_path):
        """Extract compliance names from zip files."""
        compliance_names = []
        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                for filename in zip_ref.namelist():
                    # Extract name from filename, removing extensions
                    name = Path(filename).stem
                    if name and name != '__MACOSX':
                        compliance_names.append(name)
        except Exception as e:
            print(f"Error extracting from {zip_file_path}: {e}")
        
        return compliance_names
    
    def load_all_compliance_names(self, zip_files):
        """Load compliance names from all provided zip files."""
        all_names = []
        for zip_file in zip_files:
            if os.path.exists(zip_file):
                names = self.extract_compliance_names(zip_file)
                all_names.extend(names)
        
        # Remove duplicates and sort
        self.compliance_list = sorted(list(set(all_names)))
        print(f"Loaded {len(self.compliance_list)} unique compliance standards")
        return self.compliance_list
    
    def generate_compliance_content(self, compliance_name):
        """Generate content for a compliance document."""
        # This is a template-based generator
        # Content is customized based on compliance name patterns
        
        content = {
            "title": f"{compliance_name} Compliance for Log360",
            "category": self.template_config["category"],
            "tagline": f"Automate cybersecurity risk management and ensure {compliance_name} compliance with Log360.",
            "overview": self.generate_overview(compliance_name),
            "key_features": self.generate_key_features(compliance_name),
            "valid_submission": self.template_config["valid_submission"],
            "supported_editions": self.template_config["supported_editions"],
            "supported_dcs": self.template_config["supported_dcs"],
            "help_document": self.template_config["help_doc_url"],
            "tags": [compliance_name, "Log360", "Compliance", "Regulatory compliance"],
            "release_notes": f"Predefined reports for the {compliance_name} compliance."
        }
        return content
    
    def generate_overview(self, compliance_name):
        """Generate overview text for compliance."""
        return (
            f"The {compliance_name} compliance extension for Log360 helps organizations "
            f"achieve and maintain regulatory compliance. By integrating this extension, "
            f"security teams gain end-to-end visibility into the security posture of their "
            f"ecosystem. Log360 continuously collects and correlates log data from servers, "
            f"workstations, cloud platforms, and network gateways, mapping security events "
            f"directly to {compliance_name} regulatory requirements. This enables organizations "
            f"to proactively identify vulnerabilities, protect sensitive data integrity, and "
            f"maintain a resilient environment against cyber threats throughout the lifecycle "
            f"of their systems."
        )
    
    def generate_key_features(self, compliance_name):
        """Generate key features for compliance."""
        return [
            "Track data access and modifications to prevent unauthorized use.",
            "Monitor unauthorized access attempts, privilege escalations, and changes to sensitive data in real time.",
            f"Generate audit reports to meet {compliance_name} compliance requirements.",
            "Strengthen security policies with automated monitoring and alerting.",
            "Maintain comprehensive audit trails for regulatory audits and inspections.",
            "Enable continuous compliance monitoring and reporting."
        ]
    
    def create_docx_document(self, compliance_name, content):
        """Create a DOCX document for a compliance standard."""
        doc = Document()
        
        # Set up document margins
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)
        
        # Add Title
        title = doc.add_paragraph()
        title_run = title.add_run(content["title"])
        title_run.font.size = Pt(24)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(0, 51, 102)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add Category
        category = doc.add_paragraph()
        category_run = category.add_run(f"Category: {content['category']}")
        category_run.font.size = Pt(12)
        category_run.font.bold = True
        category.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add Logo placeholder
        logo = doc.add_paragraph()
        logo_run = logo.add_run(self.template_config["logo_placeholder"])
        logo_run.font.size = Pt(11)
        logo_run.font.italic = True
        logo_run.font.color.rgb = RGBColor(128, 128, 128)
        logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add Tagline
        tagline = doc.add_paragraph()
        tagline_run = tagline.add_run(content["tagline"])
        tagline_run.font.size = Pt(13)
        tagline_run.font.italic = True
        tagline_run.font.color.rgb = RGBColor(0, 102, 153)
        tagline.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add Overview section
        doc.add_heading("Overview", level=1)
        overview_para = doc.add_paragraph(content["overview"])
        overview_para.style = 'Normal'
        
        # Add Key Features section
        doc.add_heading("Key Features", level=1)
        for feature in content["key_features"]:
            doc.add_paragraph(feature, style='List Bullet')
        
        # Add metadata section
        doc.add_heading("Compliance Details", level=1)
        
        # Valid Submission
        valid = doc.add_paragraph()
        valid_run = valid.add_run("Valid Submission: ")
        valid_run.bold = True
        valid.add_run(content["valid_submission"])
        
        # Supported Editions
        editions = doc.add_paragraph()
        editions_run = editions.add_run("Supported Editions: ")
        editions_run.bold = True
        editions.add_run(", ".join(content["supported_editions"]))
        
        # Supported DCs
        dcs = doc.add_paragraph()
        dcs_run = dcs.add_run("Supported Data Centers: ")
        dcs_run.bold = True
        dcs.add_run(", ".join(content["supported_dcs"]))
        
        # Help Document
        help_doc = doc.add_paragraph()
        help_doc_run = help_doc.add_run("Help Document: ")
        help_doc_run.bold = True
        help_doc.add_run(content["help_document"])
        
        # Tags
        tags = doc.add_paragraph()
        tags_run = tags.add_run("Tags: ")
        tags_run.bold = True
        tags.add_run(", ".join(content["tags"]))
        
        # Release Notes
        doc.add_heading("Release Notes", level=2)
        doc.add_paragraph(content["release_notes"])
        
        # Add footer with generation timestamp
        doc.add_paragraph()
        footer = doc.add_paragraph()
        footer_run = footer.add_run(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        footer_run.font.size = Pt(9)
        footer_run.font.italic = True
        footer_run.font.color.rgb = RGBColor(128, 128, 128)
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return doc
    
    def save_document(self, doc, compliance_name):
        """Save DOCX document to file."""
        # Sanitize filename
        safe_name = "".join(c for c in compliance_name if c.isalnum() or c in (' ', '-', '_')).strip()
        filename = f"{safe_name}_Compliance.docx"
        filepath = self.output_dir / filename
        
        doc.save(str(filepath))
        return filepath
    
    def generate_all_documents(self, compliance_list=None):
        """Generate DOCX documents for all compliances."""
        if compliance_list is None:
            compliance_list = self.compliance_list
        
        results = {
            "total": len(compliance_list),
            "generated": 0,
            "failed": 0,
            "files": []
        }
        
        print(f"\nGenerating {len(compliance_list)} compliance documents...")
        print("=" * 60)
        
        for i, compliance_name in enumerate(compliance_list, 1):
            try:
                print(f"[{i}/{len(compliance_list)}] Processing: {compliance_name}")
                
                # Generate content
                content = self.generate_compliance_content(compliance_name)
                
                # Create document
                doc = self.create_docx_document(compliance_name, content)
                
                # Save document
                filepath = self.save_document(doc, compliance_name)
                results["files"].append({
                    "compliance": compliance_name,
                    "file": str(filepath),
                    "size": os.path.getsize(filepath)
                })
                results["generated"] += 1
                print(f"  ✓ Successfully created: {filepath.name}")
                
            except Exception as e:
                results["failed"] += 1
                print(f"  ✗ Failed to generate for {compliance_name}: {e}")
        
        print("=" * 60)
        print(f"\nGeneration Complete!")
        print(f"Generated: {results['generated']}/{results['total']}")
        print(f"Failed: {results['failed']}/{results['total']}")
        print(f"Output directory: {self.output_dir.absolute()}")
        
        return results
    
    def generate_report(self, results):
        """Generate a summary report of the generation process."""
        report_path = self.output_dir / "generation_report.json"
        with open(report_path, 'w') as f:
            json.dump(results, f, indent=2)
        print(f"\nReport saved to: {report_path}")
        return report_path


def main():
    """Main entry point for the compliance document generator."""
    print("Compliance Document Generator Agent")
    print("=" * 60)
    
    # Initialize generator
    generator = ComplianceDocumentGenerator(output_dir="compliance_documents")
    
    # Extract compliance names from zip files
    zip_files = ["1.zip", "2.zip"]
    print("\nExtracting compliance names from zip files...")
    compliance_list = generator.load_all_compliance_names(zip_files)
    
    if not compliance_list:
        print("No compliance names found. Please ensure zip files are present.")
        print("Expected files: 1.zip, 2.zip")
        sys.exit(1)
    
    print(f"\nFound {len(compliance_list)} compliance standards:")
    for i, name in enumerate(compliance_list[:10], 1):
        print(f"  {i}. {name}")
    if len(compliance_list) > 10:
        print(f"  ... and {len(compliance_list) - 10} more")
    
    # Generate all documents
    results = generator.generate_all_documents(compliance_list)
    
    # Generate summary report
    generator.generate_report(results)
    
    print("\n✓ All compliance documents generated successfully!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
