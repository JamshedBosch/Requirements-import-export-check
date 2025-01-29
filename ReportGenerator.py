from datetime import datetime
import os
from typing import Dict, Any, List
import difflib


class ReportGenerator:
    """Generates reports from validation findings."""

    @staticmethod
    def generate_report_old(file_path, report_folder, findings):
        """Generate a structured and flexible text report for findings."""
        report_file = os.path.join(report_folder,
                                   f"{os.path.basename(file_path).replace('.xlsx', '')}_report.txt")

        with open(report_file, 'w') as f:
            # Header
            f.write(f"#### Report Summary\n")
            f.write(f"**File Name:** `{os.path.basename(file_path)}`\n")
            f.write(f"**Total Findings:** {len(findings)}\n")
            f.write("\n---\n\n")

            # Detailed Issues
            if findings:
                f.write("### Detailed Issues\n")
                for idx, finding in enumerate(findings, start=1):
                    f.write(f"#### Finding {idx}\n")

                    # Dynamically iterate over all keys in the finding dictionary
                    for key, value in finding.items():
                        # Format multi-line text blocks (like Object Text)
                        if isinstance(value, str) and "\n" in value:
                            f.write(f"- **{key}:**\n")
                            f.write(f"  ```\n{value}\n  ```\n")
                        else:
                            f.write(f"- **{key}:** `{value}`\n")

                    f.write("\n---\n\n")  # Separator between issues
            else:
                f.write("No issues found.\n")

        return report_file

    @staticmethod
    def get_html_style():
        """Return the CSS styles for the report."""
        return """
                   body {
                       font-family: Arial, sans-serif;
                       margin: 20px;
                       padding: 20px;
                       background: #f4f7f9;
                   }
                   .container {
                       max-width: 900px;
                       margin: auto;
                       background: #fff;
                       padding: 20px;
                       border-radius: 10px;
                       box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.1);
                   }
                   h1 {
                       color: #003366;
                   }
                   .issue {
                       border-left: 5px solid #ff6b6b;
                       padding: 15px;
                       margin-bottom: 20px;
                       background: #fff3f3;
                       border-radius: 8px;
                   }
                   .issue h2 {
                       color: #d43f3f;
                       margin: 0 0 8px;
                       font-size: 20px;
                   }
                   .issue p {
                       margin: 5px 0;
                       font-size: 14px;
                   }
                   .code-block {
                       background: #1e1e1e;
                       color: #ffffff;
                       padding: 10px;
                       border-radius: 5px;
                       font-family: 'Courier New', monospace;
                       white-space: pre-wrap;
                       overflow-x: auto;
                       border: 1px solid #ccc;
                       font-size: 13px;
                       line-height: 1.5;
                   }
                   .footer {
                       margin-top: 20px;
                       font-size: 12px;
                       color: #666;
                       text-align: center;
                   }
                   .diff-add {
                       background-color: #2da44e;
                       color: white;
                   }
                   .diff-del {
                       background-color: #cf222e;
                       color: white;
                   }
                   .text-block {
                       margin: 10px 0;
                       padding: 10px;
                       background: #2d2d2d;
                       border-radius: 5px;
                   }
               """

    @staticmethod
    def highlight_differences(text1: str, text2: str) -> tuple[str, str]:
        """
        Highlight the differences between two texts using HTML spans,
        ignoring whitespace in the comparison but preserving it in display.
        """

        def normalize_for_comparison(text):
            """Normalize text for comparison while keeping original positions"""
            char_mapping = []  # Keep track of original positions
            normalized = []
            pos = 0

            for char in text:
                if not (char.isspace() or char in [';', "'", '"']):
                    normalized.append(char)
                    char_mapping.append(pos)
                pos += 1

            return ''.join(normalized), char_mapping

        # Get normalized versions and position mappings
        norm1, map1 = normalize_for_comparison(text1)
        norm2, map2 = normalize_for_comparison(text2)

        # Find differences in normalized text
        matcher = difflib.SequenceMatcher(None, norm1, norm2)

        def build_highlighted(original_text, char_mapping, matcher, is_first):
            result = []
            current_pos = 0

            for op, i1, i2, j1, j2 in matcher.get_opcodes():
                # Handle text before the difference
                while current_pos < len(original_text) and (
                        not char_mapping or current_pos < char_mapping[
                    i1 if is_first else j1]):
                    result.append(original_text[current_pos])
                    current_pos += 1

                if op == 'equal':
                    # Add characters from original text for this range
                    while current_pos < len(original_text) and (
                            char_mapping and current_pos <= char_mapping[
                        i2 - 1 if is_first else j2 - 1]):
                        result.append(original_text[current_pos])
                        current_pos += 1
                else:
                    # Start difference span
                    result.append(
                        '<span class="diff-del">' if op != 'insert' else '<span class="diff-add">')

                    # Add characters from original text for this range
                    if op != 'insert' and is_first or op != 'delete' and not is_first:
                        end_pos = char_mapping[
                                      i2 - 1 if is_first else j2 - 1] + 1 if i2 > 0 and j2 > 0 else current_pos
                        while current_pos < len(
                                original_text) and current_pos < end_pos:
                            result.append(original_text[current_pos])
                            current_pos += 1

                    result.append('</span>')

            # Add any remaining text
            while current_pos < len(original_text):
                result.append(original_text[current_pos])
                current_pos += 1

            return ''.join(result)

        highlighted1 = build_highlighted(text1, map1, matcher, True)
        highlighted2 = build_highlighted(text2, map2, matcher, False)

        return highlighted1, highlighted2

    @staticmethod
    def generate_html_content(file_name, total_issues, issues_content):
        """Generate the complete HTML content."""
        # Truncate the file name if it's too long
        max_filename_length = 50
        truncated_filename = file_name[:max_filename_length] + "..." if len(
            file_name) > max_filename_length else file_name
        return f"""<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Analysis Report - {truncated_filename}</title>
            <style>
                {ReportGenerator.get_html_style()}
            </style>
        </head>
        <body>
            <div class="container">
                <h2>📋 Report for the File: {truncated_filename}</h2>
                <p><strong>Total Findings:</strong> {total_issues}</p>
        {issues_content}
                <div class="footer">
                    Generated by Import/Export Checker | Date: {datetime.now().strftime('%Y-%m-%d')}
                </div>
            </div>
        </body>
        </html>"""

    @staticmethod
    def format_issue(finding):
        """Format a single issue for the report."""
        # Extract customer and Bosch texts from the Value string
        value_lines = finding['Value'].split('\n')
        customer_text = None
        bosch_text = None

        for i, line in enumerate(value_lines):
            if "Customer File Object Text:" in line:
                customer_text = line.replace("Customer File Object Text:",
                                             "").strip()
            elif "Bosch File Object Text:" in line:
                bosch_text = line.replace("Bosch File Object Text:",
                                          "").strip()

        # If we found both texts, highlight their differences
        if customer_text is not None and bosch_text is not None:
            highlighted_customer, highlighted_bosch = ReportGenerator.highlight_differences(
                customer_text, bosch_text)

            # Replace the original texts in value_lines with highlighted versions
            for i, line in enumerate(value_lines):
                if "Customer File Object Text:" in line:
                    value_lines[
                        i] = f"       Customer File Object Text: {highlighted_customer}"
                elif "Bosch File Object Text:" in line:
                    value_lines[
                        i] = f"       Bosch File Object Text: {highlighted_bosch}"

        formatted_value = "<br>".join(value_lines)

        return f"""        <div class="issue">
                       <h2>⚠️ Row: {finding['Row']}</h2>
                       <p><strong>Attributes:</strong> {finding['Attribute']}</p>
                       <p><strong>Check:</strong> {finding['Issue']}</p>
                       <p><strong>Details:</strong></p>
                       <div class="code-block">{formatted_value}</div>
                   </div>"""

    @staticmethod
    def generate_report(file_path, report_folder, findings):
        """
        Generate a visually enhanced and more readable HTML report.

        Args:
            file_path (str): Path to the input Excel file
            report_folder (str): Directory where the report should be saved
            findings (list): List of dictionaries containing issue details

        Returns:
            str: Path to the generated report file
        """
        # Create report filename
        base_name = os.path.basename(file_path).replace('.xlsx', '')
        report_file = os.path.join(report_folder, f"{base_name}_report.html")

        # Generate issues content
        issues_content = "\n".join(
            ReportGenerator.format_issue(finding) for finding in findings)

        # Generate the complete HTML content
        html_content = ReportGenerator.generate_html_content(
            file_name=os.path.basename(file_path),
            total_issues=len(findings),
            issues_content=issues_content
        )

        # Write the report
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        return report_file
