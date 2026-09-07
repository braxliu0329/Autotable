import os
import sys
import aspose.words as aw

class StructuralMarkdownConverter:
    def __init__(self):
        self.markdown_lines = []
        self.table_anchors = []
        self.table_counter = 0

    def convert(self, doc):
        self.markdown_lines = []
        self.table_anchors = []
        self.table_counter = 0

        for section in doc.sections:
            section = section.as_section()
            self.process_story(section.body)
        
        return "\n".join(self.markdown_lines)

    def process_story(self, story):
        # The python wrapper for Aspose uses properties that might differ slightly or require method calls
        # In .NET it's story.ChildNodes, in Python it might be story.child_nodes but let's check exact API.
        # Actually, story inherits from CompositeNode. 
        # Using get_child_nodes(NodeType.ANY, False) is safer if property access fails.
        nodes = story.get_child_nodes(aw.NodeType.ANY, False)
        for node in nodes:
            if node.node_type == aw.NodeType.PARAGRAPH:
                self.process_paragraph(node.as_paragraph())
            elif node.node_type == aw.NodeType.TABLE:
                self.process_table(node.as_table())

    def process_paragraph(self, para):
        text = para.to_string(aw.SaveFormat.TEXT).strip()
        if text:
            self.markdown_lines.append(text + "\n")

    def process_table(self, table):
        table_index = self.table_counter
        self.table_counter += 1
        
        anchor_id = f"table-{table_index}"
        self.table_anchors.append((table_index, anchor_id))
        
        self.markdown_lines.append(f"\n<!-- Anchor: Table {table_index} -->")
        
        rows = table.rows
        
        for r_idx in range(rows.count):
            row = rows[r_idx]
            cells = row.cells
            
            line_parts = ["|"]
            
            for c_idx in range(cells.count):
                cell = cells[c_idx]
                cell_text = cell.to_string(aw.SaveFormat.TEXT).strip()
                cell_text = cell_text.replace("\r", " ").replace("\n", " ")
                
                cell_content = f"(R{r_idx},C{c_idx}) {cell_text}"
                line_parts.append(f" {cell_content} |")
            
            self.markdown_lines.append("".join(line_parts))
            
            if r_idx == 0:
                sep_parts = ["|"]
                for i in range(cells.count):
                    sep_parts.append(" --- |")
                self.markdown_lines.append("".join(sep_parts))
        
        self.markdown_lines.append("")

def create_sample_document(path):
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    
    builder.writeln("Header Text")
    
    builder.start_table()
    builder.insert_cell()
    builder.write("Row 0, Col 0")
    builder.insert_cell()
    builder.write("Row 0, Col 1")
    builder.end_row()
    
    builder.insert_cell()
    builder.write("Row 1, Col 0")
    builder.insert_cell()
    builder.write("Row 1, Col 1")
    builder.end_row()
    builder.end_table()
    
    builder.writeln("Footer Text")
    
    doc.save(path)

def apply_license():
    try:
        license = aw.License()
        # Try to load license from the same directory as the script or specific path
        # Assuming the license file might be in the dotnet project bin folder or root
        license_path = os.path.join(os.path.dirname(__file__), "Aspose.Total.NET.lic")
        
        # Also check the dotnet bin path where we saw it earlier
        if not os.path.exists(license_path):
             dotnet_lic_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "StructuralAnchoring", "bin", "Debug", "net9.0", "Aspose.Total.NET.lic"))
             if os.path.exists(dotnet_lic_path):
                 license_path = dotnet_lic_path

        if os.path.exists(license_path):
            license.set_license(license_path)
            print(f"INFO: Aspose license applied from {license_path}")
        else:
            print("INFO: No license file found, running in evaluation mode.")
            
    except Exception as ex:
        print(f"WARNING: Failed to apply Aspose license: {ex}")

def main():
    apply_license()

    input_path = "Output.docx"
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    
    if not os.path.exists(input_path):
        if input_path == "Output.docx":
            print("No input file provided and Output.docx not found. Creating a sample document...")
            create_sample_document(input_path)
        else:
            print(f"Error: File not found at {input_path}")
            sys.exit(1)
            
    print(f"Processing: {input_path}")
    
    try:
        doc = aw.Document(input_path)
        converter = StructuralMarkdownConverter()
        markdown_output = converter.convert(doc)
        
        output_filename = "structural_output.md"
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(markdown_output)
            
        print(f"Conversion complete. Saved to {output_filename}")
        
        print("Preview of Table Anchors:")
        for idx, anchor_id in converter.table_anchors:
            print(f"Index={idx}, AnchorId={anchor_id}")
            
        print("\nPreview of Markdown Content (first 500 chars):")
        preview_len = 500 if len(markdown_output) > 500 else len(markdown_output)
        print(markdown_output[:preview_len])
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
