# WordReorder - Word Document Reorganizer

A Command-Line Interface (CLI) tool to reorganize Microsoft Word (.docx) documents based on a specified Table of Contents (ToC) structure defined in a YAML file. This tool leverages Word's built-in heading styles (Heading 1, Heading 2, etc.) to identify sections.

This is a test project developed with the assistance of AI tools. It serves as an experimental implementation to demonstrate document reorganization capabilities and is not intended for production use without further testing and validation.

## Features

*   **Generate TOC Structure:** Analyzes a Word document and generates a structured YAML file representing its current heading hierarchy.
*   **Reorganize Document:** Creates a new Word document with sections rearranged according to the order specified in a YAML ToC file.
*   **Handles Hierarchy:** Understands nested heading levels (e.g., Heading 2 under Heading 1).
*   **Flexible Handling:** Provides options for dealing with sections present in the source but not the ToC (append, delete, warn) and vice-versa (error, warn, ignore).
*   **Cross-Platform:** Built with Python and libraries that work on Windows, macOS, and Linux.

## Important Notes & Limitations

*   **Style Dependency:** The accuracy of section identification relies *entirely* on the consistent use of Word's built-in heading styles (e.g., "Heading 1", "Heading 2") in the source documents. Content under paragraphs not styled as headings will belong to the preceding heading.
*   **Element Copying Fidelity:** The script attempts to deep-copy content elements (text, tables, images, etc.). While generally effective, complex formatting, embedded objects, tracked changes, or some drawing elements might not transfer perfectly. **Always review the reorganized document carefully.**
*   **Duplicate Headings:** If the source document contains identical heading text at the same level, the `reorganize` command currently uses the *first* section encountered with that text and warns about duplicates. The `generate` command will list all occurrences found. Ensure unique headings where structure depends on it.
*   **Headers, Footers, Footnotes:** This tool primarily reorganizes the main body content based on heading styles. It does not explicitly manage or reorder content within headers, footers, footnotes, or endnotes. These elements are generally preserved by `python-docx` but their association might change based on how content flows across pages in the new structure.
*   **Performance:** Very large documents may take some time to parse and rewrite.
