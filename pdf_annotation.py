#### Import necessary libraries ####
import fitz
import os
import argparse
import textwrap  # For text wrapping


#### Parameter settings for pdf annotation and indexing ####

# Default configurations (can be changed via function parameters in CLI)
ANNOTATED_FOLDER = "annotated_documents"
RECTANGLE_LINE_WIDTH = 0.5  # Default line width for rectangles
RECTANGLE_COLOR = (0, 1, 0)  # Green -> RGB values
INDEX_COLOR = (1, 0, 0)  # Red
INDEX_FONT_SIZE = 8  # Default font size for index annotations
INDEX_OFFSET = (5, -5)  # Offset for placing the index as a superscript
TABLE_FONT_SIZE = 8  # Font size for the index table
LINE_SPACING = 12  # Default line spacing for index table
MAX_ENTRIES_PER_PAGE = 50  # Maximum entries before creating a new index page


#### Text wrapping and Index table generation ####

# Function to wrap text to fit within the table width limit
def wrap_text(text, max_width=80):
    """Wrap text to fit within the table width limit."""
    return textwrap.fill(text, width=max_width)

# Function to add an index table to the PDF
def add_index_table(doc, indexed_texts, table_font_size, line_spacing, max_entries_per_page):
    """Adds a structured index table to the beginning of the PDF with indexed text and locations."""

    if not indexed_texts:
        print("⚠️ No indexed text found. Skipping index table.")
        return

    # Define page dimensions
    page_width = 595  # A4 width in points
    margin = 50
    max_width = page_width - 2 * margin  # Max width for text
    y_position = 80  # Initial Y position for index table

    # Create a new first page for the index table
    index_page = doc.new_page(0)
    index_page.insert_text((margin, 50), "📌 Text Index Table", fontsize=12, color=(0, 0, 0))

    entry_count = 0

    for i, (index, text, page_num) in enumerate(indexed_texts):
        # Auto-create a new index page if we exceed the limit
        if entry_count >= max_entries_per_page or y_position > 800:
            index_page = doc.new_page(0)
            index_page.insert_text((margin, 50), "📌 Text Index Table (Continued)", fontsize=12, color=(0, 0, 0))
            y_position = 80  # Reset Y position
            entry_count = 0  # Reset entry count

        # Wrap long text properly before inserting
        wrapped_text = textwrap.fill(text, width=80)

        # Insert text ensuring no overlap
        index_page.insert_text((margin, y_position), f"{index}: [Page {page_num}] {wrapped_text}",
                               fontsize=table_font_size, color=(0, 0, 0))
        y_position += line_spacing * (wrapped_text.count("\n") + 1)  # Adjust line spacing dynamically
        entry_count += 1  # Increment entry count


#### CLI arguments for parameter customization ####
def parse_cli_arguments():
    """Parse command-line arguments for user-defined customization."""
    parser = argparse.ArgumentParser(description="Annotate PDFs with indexed text regions.")

    # Output folder
    parser.add_argument("--output_folder", type=str, default="annotated_documents",
                        help="Folder to save annotated PDFs")

    # Colors
    parser.add_argument("--rectangle_color", type=float, nargs=3, default=[0, 1, 0],
                        help="Color of bounding boxes (Green default)")
    parser.add_argument("--index_color", type=float, nargs=3, default=[1, 0, 0],
                        help="Color of index numbers (Red default)")

    # Font and spacing
    parser.add_argument("--rectangle_line_width", type=float, default=0.5, help="Line width for rectangles")    
    parser.add_argument("--index_font_size", type=int, default=8, help="Font size for index annotations")
    parser.add_argument("--index_offset", type=int, nargs=2, default=[5, -5],
                        help="Offset for index numbers relative to text")
    parser.add_argument("--table_font_size", type=int, default=8, help="Font size for the index table")
    parser.add_argument("--line_spacing", type=int, default=12, help="Spacing between index table entries")
    parser.add_argument("--max_entries_per_page", type=int, default=50,
                        help="Maximum entries before creating a new index page")

    return parser.parse_args()


#### Annotating PDFs with indexed text regions ####
def annotate_pdf(file_path, output_folder, rectangle_line_width, rectangle_color, index_color, index_font_size, index_offset, 
                 table_font_size, line_spacing, max_entries_per_page):
    """
    Annotate a PDF by highlighting text regions and numbering them line by line.
    Allows customization of colors, font sizes, and positioning via CLI arguments.
    """

    # Create output folder if not exists
    os.makedirs(output_folder, exist_ok=True)

    # Load the PDF
    doc = fitz.open(file_path)
    indexed_texts = []

    for page_num, page in enumerate(doc):
        # Extract text line by line
        text_instances = page.get_text("text").splitlines()
        
        for index, line in enumerate(text_instances):
            if line.strip():  # Ensure the text is not empty
                search_instances = page.search_for(line)  # Finds bounding boxes for the text
                
                if search_instances:  # If bounding boxes exist
                    rect = search_instances[0]  # Use the first match
                    x0, y0, x1, y1 = rect  # Unpack coordinates

                    # Draw a green rectangle around the text
                    page.draw_rect(rect, color=rectangle_color, width=rectangle_line_width)

                    # Add red superscript index near the text
                    page.insert_text((x1 + index_offset[0], y0 + index_offset[1]),
                                     str(index),  # Index number
                                     fontsize=index_font_size,
                                     color=index_color)

                    # Store indexed text and its location
                    indexed_texts.append((index, line.strip(), page_num + 1))

    # Add index table pages
    add_index_table(doc, indexed_texts, table_font_size, line_spacing, max_entries_per_page)

    # Change filename to include `_annotated` suffix
    base_name = os.path.basename(file_path)  # Extract filename, e.g., "sample_invoice_1.pdf"
    name_without_ext, ext = os.path.splitext(base_name)  # Split into "sample_invoice_1" and ".pdf"
    annotated_filename = f"{name_without_ext}_annotated{ext}"  # New name: "sample_invoice_1_annotated.pdf"

    # Save the annotated PDF in the output folder
    output_file = os.path.join(output_folder, annotated_filename)
    doc.save(output_file)
    doc.close()

    print(f"✅ Annotated PDF saved: {output_file}")


#### Main function to process all PDFs in the directory ####
if __name__ == "__main__":
    args = parse_cli_arguments()
    pdf_dir = "."
    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")]

    for pdf_file in pdf_files:
        print(f"📌 Processing {pdf_file}...")
        annotate_pdf(
            pdf_file,
            output_folder=args.output_folder,
            rectangle_line_width=args.rectangle_line_width,
            rectangle_color=tuple(args.rectangle_color),
            index_color=tuple(args.index_color),
            index_font_size=args.index_font_size,
            index_offset=tuple(args.index_offset),
            table_font_size=args.table_font_size,
            line_spacing=args.line_spacing,
            max_entries_per_page=args.max_entries_per_page
        )