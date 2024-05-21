from docx import Document

def replace_dashes_with_asterisks(docx_file):
    document = Document(docx_file)
    for paragraph in document.paragraphs:
        # Split the paragraph text by newline characters
        lines = paragraph.text.split('\n')
        for i, line in enumerate(lines):
            # Search for dashes from 20 down to 3
            for dash_count in range(20, 2, -1):
                # Construct the dash sequence
                dash_sequence = '-' * dash_count
                # Check if the dash sequence is present in the line
                if dash_sequence in line:
                    # Replace the dash sequence with three asterisks
                    line = line.replace(dash_sequence, '***')
                    # Update the line in the list of lines
                    lines[i] = line
                    # Break the loop as we've found and replaced the sequence
                    break
        # Join the modified lines back into a paragraph
        modified_paragraph_text = '\n'.join(lines)
        # Clear the original paragraph and add the modified text
        paragraph.clear()
        paragraph.add_run(modified_paragraph_text)
    
    # Save the modified document
    modified_file = docx_file.replace('.docx', '_modified.docx')
    document.save(modified_file)
    print(f"Modified document saved as: {modified_file}")

# Replace dashes with asterisks in the document
replace_dashes_with_asterisks('your_document.docx')
