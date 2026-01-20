"use client";

import * as React from "react";
import * as pdfjs from "pdfjs-dist";
import { PDFDocument } from "pdf-lib";
import * as XLSX from "xlsx";
import {
  UploadCloud,
  FileText,
  FileSpreadsheet,
  X,
  Loader2,
  Download,
  PartyPopper,
  ArrowRight,
  RefreshCw,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { useToast } from "@/hooks/use-toast";
import { extractEbelnFromPdfPages } from "@/ai/flows/extract-ebeln-from-pdf-pages";
import { cn } from "@/lib/utils";

type Status = "idle" | "processing" | "success" | "error";

interface FileInputProps {
  file: File | null;
  onFileChange: (file: File | null) => void;
  placeholder: string;
  accept: string;
  icon: React.ReactNode;
}

const FileInput: React.FC<FileInputProps> = ({
  file,
  onFileChange,
  placeholder,
  accept,
  icon,
}) => {
  const inputRef = React.useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0] || null;
    onFileChange(selectedFile);
  };

  const handleRemoveFile = (e: React.MouseEvent<HTMLButtonElement>) => {
    e.stopPropagation();
    onFileChange(null);
    if (inputRef.current) {
      inputRef.current.value = "";
    }
  };

  return (
    <div
      className={cn(
        "relative flex flex-col items-center justify-center w-full p-8 border-2 border-dashed rounded-lg cursor-pointer transition-colors",
        file ? "border-primary" : "border-border hover:border-primary/50 bg-card"
      )}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={handleFileChange}
      />
      {file ? (
        <div className="text-center">
          <div className="text-primary">{icon}</div>
          <p className="mt-2 font-semibold text-foreground">{file.name}</p>
          <p className="text-xs text-muted-foreground">
            {Math.round(file.size / 1024)} KB
          </p>
          <Button
            variant="ghost"
            size="sm"
            className="mt-2 text-destructive hover:text-destructive h-auto px-2 py-1"
            onClick={handleRemoveFile}
          >
            <X className="mr-1 h-3 w-3" /> Remove
          </Button>
        </div>
      ) : (
        <div className="text-center text-muted-foreground">
          <UploadCloud className="mx-auto h-12 w-12" />
          <p className="mt-2 font-semibold">{placeholder}</p>
          <p className="text-xs">{accept.replaceAll(",", ", ")}</p>
        </div>
      )}
    </div>
  );
};

export function PdfReorderForm() {
  const [pdfFile, setPdfFile] = React.useState<File | null>(null);
  const [excelFile, setExcelFile] = React.useState<File | null>(null);
  const [status, setStatus] = React.useState<Status>("idle");
  const [progressMessage, setProgressMessage] = React.useState("");
  const [downloadUrl, setDownloadUrl] = React.useState<string | null>(null);
  const { toast } = useToast();

  React.useEffect(() => {
    // Configure PDF.js worker
    pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;
  }, []);

  const resetState = () => {
    setPdfFile(null);
    setExcelFile(null);
    setStatus("idle");
    setProgressMessage("");
    setDownloadUrl(null);
  };

  const handleReorder = async () => {
    if (!pdfFile || !excelFile) {
      toast({
        variant: "destructive",
        title: "Missing Files",
        description: "Please upload both a PDF and an Excel file.",
      });
      return;
    }

    setStatus("processing");

    try {
      // Step 1: Read files into memory
      setProgressMessage("Reading files...");
      const pdfBuffer = await pdfFile.arrayBuffer();
      const excelBuffer = await excelFile.arrayBuffer();

      // Step 2: Process Excel to get the desired order
      const getOrderedRows = async () => {
        setProgressMessage("Parsing Excel file...");
        const workbook = XLSX.read(excelBuffer, { type: "buffer" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        // Expects column A "Etiquetas de fila" to be the document number and column B "EBELN" to be the purchase order.
        const data: { "Etiquetas de fila"?: string | number; EBELN?: string | number }[] =
          XLSX.utils.sheet_to_json(worksheet);

        const filteredAndTyped = data
          .filter(row => row['Etiquetas de fila'] && row.EBELN)
          .map(row => ({
            docNumber: String(row['Etiquetas de fila']).trim(),
            ebeln: String(row.EBELN).trim(),
          }));

        // Sort by document number (column A)
        filteredAndTyped.sort((a, b) => a.docNumber.localeCompare(b.docNumber, undefined, { numeric: true }));
        
        // Return the full sorted list of objects
        return filteredAndTyped;
      };

      // Step 3: Process PDF to extract EBELN (purchase order number) from each page using GenAI
      const getEbelnToPageMap = async () => {
        setProgressMessage("Extracting text from PDF...");
        const pdfDoc = await pdfjs.getDocument(pdfBuffer.slice(0)).promise;
        const numPages = pdfDoc.numPages;
        const pageTexts: { pageNumber: number; pageText: string }[] = [];

        for (let i = 1; i <= numPages; i++) {
          const page = await pdfDoc.getPage(i);
          const textContent = await page.getTextContent();
          const pageText = textContent.items
            .map((item) => ("str" in item ? item.str : ""))
            .join(" ");
          pageTexts.push({ pageNumber: i, pageText });
        }

        setProgressMessage("Analyzing PDF pages with AI...");
        const ebelnPages = await extractEbelnFromPdfPages({ pdfPages: pageTexts });

        // A single purchase order (EBELN) can span multiple pages. 
        // This logic now correctly groups pages, assuming that a page without an EBELN
        // belongs to the preceding page that had one.
        const ebelnToPageMap = new Map<string, number[]>();
        let previousEbeln: string | null = null;
        const sortedEbelnPages = [...ebelnPages].sort((a, b) => a.pageNumber - b.pageNumber);

        for (const item of sortedEbelnPages) {
          let currentEbeln = item.ebeln ? item.ebeln.trim() : null;

          // If no EBELN is found on the current page, assume it belongs to the previous document.
          if (!currentEbeln) {
            currentEbeln = previousEbeln;
          }

          if (currentEbeln) {
            if (!ebelnToPageMap.has(currentEbeln)) {
              ebelnToPageMap.set(currentEbeln, []);
            }
            ebelnToPageMap.get(currentEbeln)!.push(item.pageNumber);
            previousEbeln = currentEbeln;
          } else {
            // Reset if a page has no EBELN and there was no previous one to associate with.
            previousEbeln = null; 
          }
        }
        
        // Sort pages for each purchase order to maintain natural order.
        for (const pages of ebelnToPageMap.values()) {
          pages.sort((a, b) => a - b);
        }

        return { ebelnToPageMap, totalPages: numPages };
      };

      const [orderedRows, { ebelnToPageMap, totalPages }] = await Promise.all([
        getOrderedRows(),
        getEbelnToPageMap(),
      ]);

      // Step 4: Determine the new page order
      setProgressMessage("Reordering pages...");
      const newPageOrder: number[] = [];
      const foundPages = new Set<number>();

      for (const row of orderedRows) {
        const ebeln = row.ebeln;
        if (ebelnToPageMap.has(ebeln)) {
          const pageNumbers = ebelnToPageMap.get(ebeln)!;
          for (const pageNumber of pageNumbers) {
            if (!foundPages.has(pageNumber)) {
              newPageOrder.push(pageNumber);
              foundPages.add(pageNumber);
            }
          }
        }
      }
      
      const originalPages = Array.from({ length: totalPages }, (_, i) => i + 1);
      const remainingPages = originalPages.filter(p => !foundPages.has(p)).sort((a, b) => a - b);
      const finalPageOrder = [...newPageOrder, ...remainingPages];


      // Step 5: Create the new PDF
      setProgressMessage("Generating new PDF...");
      const originalPdfDoc = await PDFDocument.load(pdfBuffer);
      const newPdfDoc = await PDFDocument.create();

      const copiedPageIndices = finalPageOrder.map(p => p - 1);
      const copiedPages = await newPdfDoc.copyPages(originalPdfDoc, copiedPageIndices);
      copiedPages.forEach((page) => newPdfDoc.addPage(page));

      const newPdfBytes = await newPdfDoc.save();

      // Step 6: Create download link
      const blob = new Blob([newPdfBytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setStatus("success");
    } catch (error) {
      console.error(error);
      setStatus("error");
      toast({
        variant: "destructive",
        title: "An Error Occurred",
        description:
          error instanceof Error ? error.message : "Failed to reorder PDF. Please check your files and try again.",
      });
    }
  };

  const isIdle = status === 'idle';
  const isProcessing = status === 'processing';
  const isSuccess = status === 'success';
  const isError = status === 'error';

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <FileInput
          file={pdfFile}
          onFileChange={setPdfFile}
          placeholder="Upload PDF Document"
          accept=".pdf"
          icon={<FileText className="h-12 w-12" />}
        />
        <FileInput
          file={excelFile}
          onFileChange={setExcelFile}
          placeholder="Upload Excel Sheet"
          accept=".xlsx, .xls"
          icon={<FileSpreadsheet className="h-12 w-12" />}
        />
      </div>

      <div className="flex justify-center">
        {isIdle && (
          <Button
            size="lg"
            onClick={handleReorder}
            disabled={!pdfFile || !excelFile}
          >
            Reorder PDF <ArrowRight className="ml-2 h-4 w-4" />
          </Button>
        )}

        {isProcessing && (
          <div className="flex items-center text-lg font-semibold text-primary">
            <Loader2 className="mr-3 h-6 w-6 animate-spin" />
            {progressMessage}
          </div>
        )}

        {isSuccess && (
          <div className="text-center space-y-4">
            <PartyPopper className="mx-auto h-12 w-12 text-green-500" />
            <h3 className="text-2xl font-bold">Reordering Complete!</h3>
            <p className="text-muted-foreground">Your new PDF is ready for download.</p>
            <div className="flex justify-center gap-4">
               <Button size="lg" asChild>
                <a href={downloadUrl!} download="reordered_document.pdf">
                  <Download className="mr-2 h-4 w-4" />
                  Download PDF
                </a>
              </Button>
              <Button size="lg" variant="outline" onClick={resetState}>
                Start Over
              </Button>
            </div>
          </div>
        )}

        {isError && (
          <div className="text-center space-y-4">
             <h3 className="text-2xl font-bold text-destructive">Oops! Something went wrong.</h3>
             <p className="text-muted-foreground">We couldn't reorder your PDF. Please try again.</p>
             <Button size="lg" onClick={resetState}>
               <RefreshCw className="mr-2 h-4 w-4" />
               Try Again
             </Button>
          </div>
        )}
      </div>
    </div>
  );
}
