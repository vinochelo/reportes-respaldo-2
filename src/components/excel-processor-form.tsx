"use client";

import * as React from "react";
import * as XLSX from "xlsx";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import {
  UploadCloud,
  FileSpreadsheet,
  X,
  Loader2,
  Download,
  PartyPopper,
  ArrowRight,
  RefreshCw,
  HelpCircle,
  FileText,
} from "lucide-react";

import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { Button } from "@/components/ui/button";
import { useToast } from "@/hooks/use-toast";
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

  React.useEffect(() => {
    if (file === null && inputRef.current) {
      inputRef.current.value = "";
    }
  }, [file]);

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

export function ExcelProcessorForm() {
  const [comprasFile, setComprasFile] = React.useState<File | null>(null);
  const [documentosFile, setDocumentosFile] = React.useState<File | null>(null);
  const [status, setStatus] = React.useState<Status>("idle");
  const [progressMessage, setProgressMessage] = React.useState("");
  const [downloadUrl, setDownloadUrl] = React.useState<string | null>(null);
  const { toast } = useToast();

  const resetState = () => {
    setComprasFile(null);
    setDocumentosFile(null);
    setStatus("idle");
    setProgressMessage("");
    setDownloadUrl(null);
  };
  
  const handleProcess = async () => {
    if (!comprasFile || !documentosFile) {
      toast({
        variant: "destructive",
        title: "Archivos Faltantes",
        description: "Por favor, sube ambos reportes.",
      });
      return;
    }

    setStatus("processing");

    try {
      // 1. Process "Reporte de Compras"
      setProgressMessage("Leyendo reporte de compras...");
      const comprasBuffer = await comprasFile.arrayBuffer();
      const comprasWorkbook = XLSX.read(comprasBuffer, { type: "buffer" });
      const comprasSheetName = comprasWorkbook.SheetNames[0];
      const comprasWorksheet = comprasWorkbook.Sheets[comprasSheetName];
      const comprasData: any[][] = XLSX.utils.sheet_to_json(comprasWorksheet, { header: 1 });
      
      const comprasHeaderRowIndex = comprasData.findIndex(row => 
        Array.isArray(row) && 
        row.some(cell => typeof cell === 'string' && cell.trim() === "Ord. de Compra")
      );
      if (comprasHeaderRowIndex === -1) {
        throw new Error("No se encontró la fila de encabezado con 'Ord. de Compra' en el reporte de compras.");
      }
      const comprasHeaders = comprasData[comprasHeaderRowIndex];
      const purchaseOrderColIndex = comprasHeaders.findIndex(cell => typeof cell === 'string' && cell.trim() === "Ord. de Compra");
      if (purchaseOrderColIndex === -1) {
        throw new Error("No se pudo encontrar la columna 'Ord. de Compra' en la fila de encabezado del reporte de compras.");
      }
      
      const comprasDataRows = comprasData.slice(comprasHeaderRowIndex + 1);
      const groupedByPurchaseOrder = comprasDataRows.reduce((acc, row) => {
        const poNumber = row[purchaseOrderColIndex];
        if (poNumber) {
          const poString = String(poNumber);
          if (!acc[poString]) {
            acc[poString] = [];
          }
          acc[poString].push(row);
        }
        return acc;
      }, {} as Record<string, any[][]>);

      // 2. Process "Reporte de Documentos" (EKBE)
      setProgressMessage("Leyendo reporte de documentos...");
      const docBuffer = await documentosFile.arrayBuffer();
      const docWorkbook = XLSX.read(docBuffer); // Let xlsx handle format (HTML table in XLS)
      const docSheetName = docWorkbook.SheetNames[0];
      const docWorksheet = docWorkbook.Sheets[docSheetName];
      const docData: any[][] = XLSX.utils.sheet_to_json(docWorksheet, { header: 1 });

      const docHeaderRowIndex = docData.findIndex(row => 
        Array.isArray(row) && 
        row.some(cell => typeof cell === 'string' && cell.trim().toUpperCase() === 'EBELN')
      );
      if (docHeaderRowIndex === -1) {
          throw new Error("No se encontró la fila de encabezado con 'EBELN' en el archivo de documentos.");
      }
      const docHeaders = docData[docHeaderRowIndex].map(h => String(h || '').trim().toUpperCase());
      const ebelnColIndex = docHeaders.indexOf("EBELN");
      const belnrColIndex = docHeaders.indexOf("BELNR");

      if (ebelnColIndex === -1 || belnrColIndex === -1) {
          throw new Error("No se encontraron las columnas 'EBELN' o 'BELNR' en el archivo de documentos.");
      }

      const belnrToEbelnMap = new Map<string, string>();
      const docDataRows = docData.slice(docHeaderRowIndex + 1);
      for (const row of docDataRows) {
          const ebeln = row[ebelnColIndex];
          const belnr = row[belnrColIndex];
          if (ebeln && belnr) {
              belnrToEbelnMap.set(String(belnr), String(ebeln));
          }
      }

      if (belnrToEbelnMap.size === 0) {
        throw new Error("No se encontraron documentos para procesar en el archivo de documentos.");
      }

      // 3. Generate PDF
      setProgressMessage("Generando PDF...");
      const pdfDoc = await PDFDocument.create();
      const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const helveticaBoldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      
      const sortedBelnrs = Array.from(belnrToEbelnMap.keys()).sort();
      let currentPage = pdfDoc.addPage();
      const { width, height } = currentPage.getSize();
      const margin = 40;
      let y = height - margin;

      const drawTable = (page: any, headers: string[], rows: any[][], startY: number) => {
        let currentY = startY;
        const rowHeight = 18;
        const headerSize = 9;
        const rowSize = 8;
        const colWidths = headers.map(() => (width - 2 * margin) / headers.length);
        
        // Draw Header
        let currentX = margin;
        headers.forEach((header, i) => {
          page.drawText(String(header || ''), { x: currentX + 3, y: currentY - 12, font: helveticaBoldFont, size: headerSize });
          currentX += colWidths[i];
        });
        currentY -= rowHeight;
        page.drawLine({ start: { x: margin, y: currentY + 4 }, end: { x: width - margin, y: currentY + 4 }, thickness: 0.5 });
        
        // Draw Rows
        for (const row of rows) {
            if (currentY < margin + rowHeight) {
                page = pdfDoc.addPage();
                currentY = height - margin;
                // Redraw header on new page
                let newX = margin;
                headers.forEach((header, i) => {
                  page.drawText(String(header || ''), { x: newX + 3, y: currentY - 12, font: helveticaBoldFont, size: headerSize });
                  newX += colWidths[i];
                });
                currentY -= rowHeight;
                page.drawLine({ start: { x: margin, y: currentY + 4 }, end: { x: width - margin, y: currentY + 4 }, thickness: 0.5 });
            }
            
            let cellX = margin;
            row.forEach((cell, i) => {
                page.drawText(String(cell || ''), { x: cellX + 3, y: currentY - 12, font: helveticaFont, size: rowSize, color: rgb(0.2, 0.2, 0.2) });
                cellX += colWidths[i];
            });
            currentY -= rowHeight;
        }
        return { finalY: currentY, finalPage: page };
      };

      for (const belnr of sortedBelnrs) {
        const ebeln = belnrToEbelnMap.get(belnr);
        if (!ebeln) continue;
        const poData = groupedByPurchaseOrder[ebeln];
        if (!poData) continue;

        const tableHeight = (poData.length + 1) * 18 + 50; // estimate
        if (y < margin + tableHeight) {
            currentPage = pdfDoc.addPage();
            y = height - margin;
        }

        y -= 20;
        currentPage.drawText(`Documento: ${belnr}`, { x: margin, y, font: helveticaBoldFont, size: 14, color: rgb(0,0,0) });
        y -= 15;
        currentPage.drawText(`Orden de Compra: ${ebeln}`, { x: margin, y, font: helveticaFont, size: 11, color: rgb(0.3, 0.3, 0.3) });
        y -= 25;

        const { finalY, finalPage } = drawTable(currentPage, comprasHeaders, poData, y);
        currentPage = finalPage;
        y = finalY;
        y -= 20; // Space between tables
      }

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: "application/pdf" });
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
          error instanceof Error ? error.message : "No se pudo procesar los archivos. Por favor, revisa que el formato sea correcto.",
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
          file={comprasFile}
          onFileChange={setComprasFile}
          placeholder="Subir Reporte de Compras"
          accept=".xlsx, .xls"
          icon={<FileSpreadsheet className="h-12 w-12" />}
        />
        <FileInput
          file={documentosFile}
          onFileChange={setDocumentosFile}
          placeholder="Subir Reporte Documentos (EKBE)"
          accept=".xlsx, .xls"
          icon={<FileText className="h-12 w-12" />}
        />
      </div>

      <div className="flex justify-center">
        {isIdle && (
          <Button
            size="lg"
            onClick={handleProcess}
            disabled={!comprasFile || !documentosFile}
          >
            Generar PDF <ArrowRight className="ml-2 h-4 w-4" />
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
            <h3 className="text-2xl font-bold">¡Procesamiento Completo!</h3>
            <p className="text-muted-foreground">Tu nuevo reporte en PDF está listo para descargar.</p>
            <div className="flex justify-center gap-4">
               <Button size="lg" asChild>
                <a href={downloadUrl!} download="Reporte_Consolidado.pdf">
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
             <p className="text-muted-foreground">No pudimos procesar tus reportes. Por favor, inténtalo de nuevo.</p>
             <Button size="lg" onClick={resetState}>
               <RefreshCw className="mr-2 h-4 w-4" />
               Intentar de Nuevo
             </Button>
          </div>
        )}
      </div>

      <div className="pt-4 border-t">
        <Accordion type="single" collapsible className="w-full">
          <AccordionItem value="item-1">
            <AccordionTrigger>
              <div className="flex items-center gap-2 font-semibold">
                <HelpCircle className="h-5 w-5" />
                ¿Cómo funciona el proceso?
              </div>
            </AccordionTrigger>
            <AccordionContent>
              <ol className="list-decimal space-y-2 pl-6 text-sm text-muted-foreground">
                <li>
                  <strong>Subir Reporte de Compras:</strong> Carga tu reporte principal de compras en formato Excel (.xlsx o .xls).
                </li>
                 <li>
                  <strong>Subir Reporte de Documentos:</strong> Carga tu reporte de documentos (EKBE), que puede ser un archivo HTML guardado como .xls.
                </li>
                <li>
                  <strong>Enlace de Datos:</strong> La aplicación asocia los números de documento (BELNR) del segundo archivo con sus órdenes de compra (EBELN / Ord. de Compra) correspondientes en el primer archivo.
                </li>
                <li>
                  <strong>Generación de PDF:</strong> Se crea un único documento PDF. El reporte está organizado por número de documento, y cada sección contiene la tabla de artículos de la orden de compra asociada.
                </li>
                <li>
                  <strong>Descarga:</strong> Finalmente, se te proporciona un enlace para descargar el PDF consolidado.
                </li>
              </ol>
            </AccordionContent>
          </AccordionItem>
        </Accordion>
      </div>
    </div>
  );
}
