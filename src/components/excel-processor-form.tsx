"use client";

import * as React from "react";
import * as XLSX from "xlsx";
import { PDFDocument, rgb, StandardFonts, PageSizes } from "pdf-lib";
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
      const comprasWorkbook = XLSX.read(comprasBuffer, { type: "buffer", cellDates: true });
      const comprasSheetName = comprasWorkbook.SheetNames[0];
      const comprasWorksheet = comprasWorkbook.Sheets[comprasSheetName];
      const comprasData: any[][] = XLSX.utils.sheet_to_json(comprasWorksheet, { header: 1 });
      
      const comprasHeaderRowIndex = comprasData.findIndex(row => 
        Array.isArray(row) && 
        row.some(cell => typeof cell === 'string' && cell.trim().toUpperCase() === "ORD. DE COMPRA")
      );
      if (comprasHeaderRowIndex === -1) {
        throw new Error("No se encontró la fila de encabezado con 'Ord. de Compra' en el reporte de compras.");
      }
      const comprasHeaders = comprasData[comprasHeaderRowIndex].map(h => String(h || '').trim());
      const purchaseOrderColIndex = comprasHeaders.findIndex(cell => cell.toUpperCase() === "ORD. DE COMPRA");
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
      const docWorkbook = XLSX.read(docBuffer);
      const docSheetName = docWorkbook.SheetNames[0];
      const docWorksheet = docWorkbook.Sheets[docSheetName];
      const docData: any[][] = XLSX.utils.sheet_to_json(docWorksheet, { header: 1 });

      const docHeaderRowIndex = docData.findIndex(row => {
        if (!Array.isArray(row)) return false;
        const upperCaseCells = row.map(cell => String(cell || '').trim().toUpperCase());
        return upperCaseCells.includes('EBELN') && upperCaseCells.includes('BELNR');
      });

      if (docHeaderRowIndex === -1) {
          throw new Error("No se encontró la fila de encabezado con 'EBELN' y 'BELNR' en el archivo de documentos.");
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

      const pageLayout = {
        size: [PageSizes.A4[1], PageSizes.A4[0]], // Landscape A4
        margin: 30,
      };

      const drawPageHeader = (page: any) => {
        const y = page.getSize().height - pageLayout.margin + 10;
        page.drawText("Reporte Utilidades de Pedidos de Compras", {
          x: pageLayout.margin,
          y,
          font: helveticaFont,
          size: 8,
        });
      };
      
      let currentPage = pdfDoc.addPage(pageLayout.size);
      drawPageHeader(currentPage);
      const { width, height } = currentPage.getSize();
      let y = height - pageLayout.margin - 15;

      const drawTable = (page: any, headers: string[], rows: any[][], startY: number, ebeln: string, belnr: string) => {
        let currentY = startY;
        const rowHeight = 9;
        const headerSize = 5;
        const rowSize = 5;
        const availableWidth = width - 2 * pageLayout.margin;
        
        const columnsToSum = ['Cant. In', 'Costo Uni', 'PVP S/IVA', 'Costo total', 'PVP Total', 'Valor a Pagar'];
        const columnsToAverage = ['% Utilidad'];
        const valueColumns = ['Cant. In','Costo Uni','PVP S/IVA','% Utilidad', 'Costo total', 'PVP Total', 'Valor a Pagar'];
        
        const upperHeaders = headers.map(h => String(h || '').trim().replace(/\.?$/, '').toUpperCase());
        const sumIndices = columnsToSum.map(colName => upperHeaders.indexOf(colName.toUpperCase())).filter(i => i !== -1);
        const avgIndices = columnsToAverage.map(colName => upperHeaders.indexOf(colName.toUpperCase())).filter(i => i !== -1);

        const totals = new Array(headers.length).fill(0);
        const avgTotals = new Array(headers.length).fill(0);
        const avgCounts = new Array(headers.length).fill(0);
        
        const centerAlignedIndices: number[] = [];
        headers.forEach((h, i) => {
            if (valueColumns.map(sc => sc.toUpperCase()).includes(String(h || '').trim().replace(/\.?$/, '').toUpperCase())) {
                centerAlignedIndices.push(i);
            }
        });
        
        const columnWidths = [65, 55, 270, 55, 50, 210, 45, 55, 55, 45, 50, 50, 40];
        const tableWidth = columnWidths.reduce((a, b) => a + b, 0);
        const scale = availableWidth / tableWidth;
        const scaledWidths = columnWidths.map(w => w * scale);

        // Draw Header
        const headerTopY = currentY + 2;
        page.drawRectangle({
          x: pageLayout.margin,
          y: currentY - rowHeight + 2,
          width: availableWidth,
          height: rowHeight,
          color: rgb(0.22, 0.45, 0.70), // Blue background
        })
        let currentX = pageLayout.margin;
        headers.forEach((header, i) => {
          page.drawText(String(header || ''), { x: currentX + 3, y: currentY - 4, font: helveticaBoldFont, size: headerSize, color: rgb(1,1,1) });
          currentX += scaledWidths[i];
        });
        currentY -= rowHeight;
        page.drawLine({ start: { x: pageLayout.margin, y: currentY + 2 }, end: { x: width - pageLayout.margin, y: currentY + 2 }, thickness: 0.5, color: rgb(0, 0, 0) });

        
        // Draw Rows
        for (const row of rows) {
            const rowTopY = currentY + 2;
            if (currentY < pageLayout.margin + rowHeight * 2) { // Need space for row and summary
                page = pdfDoc.addPage(pageLayout.size);
                drawPageHeader(page);
                currentY = height - pageLayout.margin - 15;

                let newX = pageLayout.margin;
                page.drawRectangle({ x: pageLayout.margin, y: currentY - rowHeight + 2, width: availableWidth, height: rowHeight, color: rgb(0.22, 0.45, 0.70) });
                headers.forEach((header, i) => {
                  page.drawText(String(header || ''), { x: newX + 3, y: currentY - 4, font: helveticaBoldFont, size: headerSize, color: rgb(1,1,1) });
                  newX += scaledWidths[i];
                });
                currentY -= rowHeight;
                page.drawLine({ start: { x: pageLayout.margin, y: currentY + 2 }, end: { x: width - pageLayout.margin, y: currentY + 2 }, thickness: 0.5, color: rgb(0, 0, 0) });
            }
            
            let cellX = pageLayout.margin;
            row.forEach((cell, i) => {
                let cellValue;
                if (cell instanceof Date) {
                    const day = String(cell.getDate()).padStart(2, '0');
                    const month = String(cell.getMonth() + 1).padStart(2, '0');
                    const year = cell.getFullYear();
                    cellValue = `${day}/${month}/${year}`;
                } else {
                    cellValue = String(cell === null || cell === undefined ? '' : cell);
                }
                
                const isCenterAligned = centerAlignedIndices.includes(i);
                const textWidth = helveticaFont.widthOfTextAtSize(cellValue, rowSize);
                let xPos = cellX + 3; // Default left alignment

                if (isCenterAligned) {
                    xPos = cellX + (scaledWidths[i] - textWidth) / 2;
                }

                page.drawText(cellValue, { x: xPos, y: currentY - 4, font: helveticaFont, size: rowSize, color: rgb(0.2, 0.2, 0.2) });

                const numValue = parseFloat(cellValue.replace(/,/g, ''));
                if (!isNaN(numValue)) {
                    if (sumIndices.includes(i)) {
                        totals[i] += numValue;
                    }
                    if (avgIndices.includes(i)) {
                        avgTotals[i] += numValue;
                        avgCounts[i]++;
                    }
                }
                cellX += scaledWidths[i];
            });
            currentY -= rowHeight;
            
            // Draw grid lines for the row
            const rowBottomY = currentY + 2;
            let vLineX = pageLayout.margin;
            for(let i=0; i <= scaledWidths.length; i++) {
                page.drawLine({ start: {x: vLineX, y: rowTopY}, end: {x: vLineX, y: rowBottomY}, color: rgb(0, 0, 0), thickness: 0.5});
                if (i < scaledWidths.length) vLineX += scaledWidths[i];
            }
        }

        // Draw summary row
        const summaryTopY = currentY + 2;
        page.drawLine({ start: { x: pageLayout.margin, y: summaryTopY }, end: { x: width - pageLayout.margin, y: summaryTopY }, thickness: 0.5, color: rgb(0, 0, 0) });
        page.drawRectangle({ x: pageLayout.margin, y: currentY - rowHeight + 2, width: availableWidth, height: rowHeight, color: rgb(1, 1, 0.8) }); // Yellow background
        let summaryX = pageLayout.margin;
        page.drawText('*', { x: summaryX + 3, y: currentY - 4, font: helveticaBoldFont, size: rowSize });
        
        headers.forEach((h, i) => {
            let textToDraw = '';
            
            if (sumIndices.includes(i)) {
                 textToDraw = totals[i].toFixed(2);
            } else if (avgIndices.includes(i) && avgCounts[i] > 0) {
                 const average = avgTotals[i] / avgCounts[i];
                 textToDraw = average.toFixed(2) + '%';
            }

            if (textToDraw) {
                const isCenterAligned = centerAlignedIndices.includes(i);
                const textWidth = helveticaBoldFont.widthOfTextAtSize(textToDraw, rowSize);
                let xPos = summaryX + 3;
                if (isCenterAligned) {
                    xPos = summaryX + (scaledWidths[i] - textWidth) / 2;
                }
                page.drawText(textToDraw, { x: xPos, y: currentY - 4, font: helveticaBoldFont, size: rowSize });
            }
            summaryX += scaledWidths[i];
        });
        currentY -= rowHeight;
        
        const summaryBottomY = currentY + 2;
        page.drawLine({ start: { x: pageLayout.margin, y: summaryBottomY }, end: { x: width - pageLayout.margin, y: summaryBottomY }, thickness: 0.5, color: rgb(0, 0, 0) });
        let vLineX = pageLayout.margin;
        for(let i=0; i <= scaledWidths.length; i++) {
            page.drawLine({ start: {x: vLineX, y: summaryTopY}, end: {x: vLineX, y: summaryBottomY}, color: rgb(0, 0, 0), thickness: 0.5});
            if (i < scaledWidths.length) vLineX += scaledWidths[i];
        }


        return { finalY: currentY, finalPage: page };
      };

      for (const [index, belnr] of sortedBelnrs.entries()) {
        const ebeln = belnrToEbelnMap.get(belnr);
        if (!ebeln) continue;
        const poData = groupedByPurchaseOrder[ebeln];
        if (!poData || poData.length === 0) continue;

        if (index > 0) {
          currentPage = pdfDoc.addPage(pageLayout.size);
          y = height - pageLayout.margin - 15;
        }
        drawPageHeader(currentPage);
        
        const { finalPage } = drawTable(currentPage, comprasHeaders, poData, y, ebeln, belnr);
        currentPage = finalPage;
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
