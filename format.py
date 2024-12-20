from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING

# Create a new Document
doc = Document()

# Add a title to the document
title = doc.add_heading('The Wonders of Nature', level=1)
title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Center the title

# Add an introductory paragraph
intro = doc.add_paragraph()
intro.add_run("Nature is a vast and breathtaking entity. ").bold = True
intro.add_run("Its beauty and complexity have inspired poets, scientists, and adventurers throughout history.")
intro_format = intro.paragraph_format
intro_format.space_after = Pt(12)  # Add some space after the paragraph

# Add a section with custom formatting
paragraph = doc.add_paragraph('Let us explore some of the key aspects of nature that make it so captivating:')
paragraph_format = paragraph.paragraph_format
paragraph_format.left_indent = Pt(24)  # Add left indent
paragraph_format.line_spacing = 1.5  # Set line spacing

# Add a bulleted list
doc.add_paragraph('The vastness of the universe', style='List Bullet')
doc.add_paragraph('The intricate designs of flowers and trees', style='List Bullet')
doc.add_paragraph('The resilience of ecosystems', style='List Bullet')

# Add another paragraph with styled text
conclusion = doc.add_paragraph()
conclusion.add_run("Nature teaches us to respect and cherish the world we live in.").italic = True

# Add a footer with page numbering
section = doc.sections[0]
footer = section.footer
footer_paragraph = footer.paragraphs[0]
footer_paragraph.text = "Page "
footer_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

# Save the document
doc.save('New_Document.docx')
