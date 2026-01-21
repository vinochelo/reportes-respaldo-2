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
import { groupPagesByEbeln } from "@/ai/flows/extract-ebeln-from-pdf-pages";
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
            <X className="mr-1 h-3 w-3" /> Quitar
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
        title: "Archivos Faltantes",
        description: "Por favor, sube un archivo PDF y un archivo de Excel.",
      });
      return;
    }

    setStatus("processing");

    try {
      // Step 1: Read files into memory
      setProgressMessage("Leyendo archivos...");
      const pdfBuffer = await pdfFile.arrayBuffer();
      const excelBuffer = await excelFile.arrayBuffer();

      // Step 2: Process Excel to get the desired order
      const getOrderedRows = async () => {
        setProgressMessage("Procesando archivo de Excel...");
        const workbook = XLSX.read(excelBuffer, { type: "buffer" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        // Convert to array of arrays, ignoring headers. This is more robust.
        const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Assumes column A is doc number, column B is EBELN. Skips header row (index 0).
        const filteredAndTyped = data
          .slice(1) // Skip header row
          .filter(row => row && row[0] && row[1]) // Ensure both columns have data
          .map(row => ({
            docNumber: String(row[0]).trim(),
            ebeln: String(row[1]).trim(),
          }));

        // Sort by document number (column A)
        filteredAndTyped.sort((a, b) => a.docNumber.localeCompare(b.docNumber, undefined, { numeric: true }));
        
        // Return the full sorted list of objects
        return filteredAndTyped;
      };

      // Step 3: Process PDF to group pages by EBELN using GenAI
      const getEbelnToPageMap = async () => {
        setProgressMessage("Extrayendo texto del PDF...");
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

        setProgressMessage("Analizando y agrupando páginas del PDF con IA...");
        const ebelnGroups = await groupPagesByEbeln({ pdfPages: pageTexts });

        const ebelnToPageMap = new Map<string, number[]>();
        for (const group of ebelnGroups) {
            const ebeln = group.ebeln.trim();
            if (ebeln) {
                // The AI returns page numbers, sort them to be safe.
                const sortedPages = group.pageNumbers.sort((a, b) => a - b);
                ebelnToPageMap.set(ebeln, sortedPages);
            }
        }

        return { ebelnToPageMap, totalPages: numPages };
      };

      const [orderedRows, { ebelnToPageMap, totalPages }] = await Promise.all([
        getOrderedRows(),
        getEbelnToPageMap(),
      ]);

      // Step 4: Determine the new page order
      setProgressMessage("Reordenando páginas...");
      const newPageOrder: number[] = [];
      const foundPages = new Set<number>();

      for (const row of orderedRows) {
        const ebeln = row.ebeln;
        if (ebelnToPagePageMap.has(ebeln)) {
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
      setProgressMessage("Generando nuevo PDF...");
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
        title: "Ocurrió un Error",
        description:
          error instanceof Error ? error.message : "No se pudo reordenar el PDF. Por favor, revisa tus archivos e inténtalo de nuevo.",
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
          placeholder="Subir Documento PDF"
          accept=".pdf"
          icon={<FileText className="h-12 w-12" />}
        />
        <FileInput
          file={excelFile}
          onFileChange={setExcelFile}
          placeholder="Subir Hoja de Excel"
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
            Reordenar PDF <ArrowRight className="ml-2 h-4 w-4" />
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
            <h3 className="text-2xl font-bold">¡Reordenamiento Completo!</h3>
            <p className="text-muted-foreground">Tu nuevo PDF está listo para descargar.</p>
            <div className="flex justify-center gap-4">
               <Button size="lg" asChild>
                <a href={downloadUrl!} download="documento_reordenado.pdf">
                  <Download className="mr-2 h-4 w-4" />
                  Descargar PDF
                </a>
              </Button>
              <Button size="lg" variant="outline" onClick={resetState}>
                Empezar de Nuevo
              </Button>
            </div>
          </div>
        )}

        {isError && (
          <div className="text-center space-y-4">
             <h3 className="text-2xl font-bold text-destructive">¡Uy! Algo salió mal.</h3>
             <p className="text-muted-foreground">No pudimos reordenar tu PDF. Por favor, inténtalo de nuevo.</p>
             <Button size="lg" onClick={resetState}>
               <RefreshCw className="mr-2 h-4 w-4" />
               Intentar de Nuevo
             </Button>
          </div>
        )}
      </div>
    </div>
  );
}
