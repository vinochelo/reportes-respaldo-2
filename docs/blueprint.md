# **App Name**: PDF ReOrder

## Core Features:

- File Upload: Allow users to upload the PDF and Excel files.
- Data Extraction: Extract data from the uploaded Excel file (EBELN and BELNR columns).
- PDF Page Extraction: Split PDF pages and extract all pages in a structured data to allow ordering later
- Page Mapping: The app extract EBELN column values of each page, in order to use it as index for future ordering, using a 'tool' from LLM that reads the table and links each record to a number in order column
- PDF Reordering: Reorder the PDF pages based on the Excel data.
- New PDF Generation: Create a new PDF with the reordered pages.
- Download: Provide a download link for the reordered PDF.

## Style Guidelines:

- Primary color: Soft blue (#A0BFE0) for a calm, professional feel.
- Background color: Very light blue (#F0F5FA) for a clean look.
- Accent color: Pale orange (#E0A8A0) for subtle emphasis.
- Body and headline font: 'Inter' sans-serif for a clear and neutral presentation.
- Clean and intuitive layout for easy navigation.
- Simple and professional icons.