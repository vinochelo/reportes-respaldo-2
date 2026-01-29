"use client";

import * as React from "react";
import * as XLSX from "xlsx";
import { PDFDocument, rgb, StandardFonts, PageSizes, PDFFont } from "pdf-lib";
import {
  UploadCloud,
  FileSpreadsheet,
  X,
  Loader2,
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
        file ? "border-green-500 bg-green-500/10" : "border-border hover:border-primary/50 bg-card"
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
          <div className="text-green-500">{icon}</div>
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

  React.useEffect(() => {
    if (status === 'success' && downloadUrl) {
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.setAttribute('download', 'Reporte_Consolidado.pdf');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      // Clean up the object URL after download
      URL.revokeObjectURL(downloadUrl);
      setDownloadUrl(null);
    }
  }, [status, downloadUrl]);

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
    const processedEbelns = new Set<string>();

    try {
      // 1. Process "Reporte de Utilidad"
      setProgressMessage("Leyendo reporte de utilidad...");
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
        throw new Error("No se encontró la fila de encabezado con 'Ord. de Compra' en el reporte de utilidad.");
      }
      const comprasHeaders = comprasData[comprasHeaderRowIndex].map(h => String(h || '').trim());
      const purchaseOrderColIndex = comprasHeaders.findIndex(cell => cell.toUpperCase() === "ORD. DE COMPRA");
      if (purchaseOrderColIndex === -1) {
        throw new Error("No se pudo encontrar la columna 'Ord. de Compra' en la fila de encabezado del reporte de utilidad.");
      }
      
      const comprasDataRows = comprasData.slice(comprasHeaderRowIndex + 1);

      const normalizePO = (po: any) => String(po).trim().replace(/^0+/, '');

      const groupedByPurchaseOrder = comprasDataRows.reduce((acc, row) => {
        const poNumber = row[purchaseOrderColIndex];
        if (poNumber) {
          const poString = normalizePO(poNumber);
          if (!acc[poString]) {
            acc[poString] = [];
          }
          acc[poString].push(row);
        }
        return acc;
      }, {} as Record<string, any[][]>);

      // 2. Process "Reporte Tabla EKBE"
      setProgressMessage("Leyendo reporte tabla EKBE...");

      let docData: any[][] = [];
      let docHeaderRowIndex = -1;
      let ebelnColIndex = -1;
      let belnrColIndex = -1;

      const findHeaders = (data: any[][]) => {
          let headerRow = -1, ebelnCol = -1, belnrCol = -1;
          for (let i = 0; i < data.length; i++) {
              const row = data[i];
              if (!Array.isArray(row) || row.length === 0) continue;

              let foundEbeln = -1;
              let foundBelnr = -1;
              
              row.forEach((cell, index) => {
                  const cellContent = String(cell || '').replace(/\.|\s/g, '').toLowerCase();
                  if (cellContent.includes('ebeln') || cellContent.includes('doccompr')) {
                      if (foundEbeln === -1) foundEbeln = index;
                  }
                  if (cellContent.includes('belnr') || cellContent.includes('docmat')) {
                      if (foundBelnr === -1) foundBelnr = index;
                  }
              });

              if (foundEbeln !== -1 && foundBelnr !== -1) {
                  headerRow = i;
                  ebelnCol = foundEbeln;
                  belnrCol = foundBelnr;
                  break; 
              }
          }
          return { headerRow, ebelnCol, belnrCol };
      };
      
      // Unified reading logic: Read as buffer and let XLSX figure it out.
      try {
          const buffer = await documentosFile.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: 'buffer' });

          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error("El archivo de documentos parece estar vacío o en un formato no reconocido por la librería.");
          }

          let bestSheetData: any[][] = [];
          let maxRows = -1;

          // Find the sheet with the most data, as SAP exports can have multiple "sheets".
          for (const sheetName of workbook.SheetNames) {
              const worksheet = workbook.Sheets[sheetName];
              const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
              if (data.length > maxRows) {
                  maxRows = data.length;
                  bestSheetData = data;
              }
          }
          
          docData = bestSheetData;

          if (docData.length === 0) {
            throw new Error("No se pudo extraer ninguna tabla de datos válida del archivo de documentos.");
          }
          
          const headers = findHeaders(docData);
          docHeaderRowIndex = headers.headerRow;
          ebelnColIndex = headers.ebelnCol;
          belnrColIndex = headers.belnrCol;

      } catch (e) {
          console.error("Error processing documents file:", e);
          const errorMessage = e instanceof Error ? e.message : String(e);
          throw new Error(`Falló la lectura del archivo de documentos. Error original: ${errorMessage}`);
      }

      // FINAL CHECK
      if (docHeaderRowIndex === -1) {
          const dataSample = (docData.length > 0 ? docData : [["No se pudo leer el archivo o está vacío."]]).slice(0, 10).map(row => JSON.stringify(row)).join('\\n');
          throw new Error(`No se encontró la fila de encabezado con ('EBELN' o 'Doc.compr.') y ('BELNR' o 'Doc.mat.') en el archivo de documentos.\nAsí es como se están leyendo las primeras 10 filas:\n${dataSample}`);
      }
      
      const belnrToEbelnMap = new Map<string, string>();
      const docDataRows = docData.slice(docHeaderRowIndex + 1);
      for (const row of docDataRows) {
          if (!row[ebelnColIndex] || !row[belnrColIndex]) continue;
          const ebeln = row[ebelnColIndex];
          const belnr = row[belnrColIndex];
          if (ebeln && belnr) {
              belnrToEbelnMap.set(String(belnr).trim(), normalizePO(ebeln));
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
      
      const COLUMN_WIDTH_CONFIG: Record<string, number> = {
        'ORD. DE COMPRA': 60,
        'PROVEEDOR': 60,
        'NOMBRE PROVEEDOR': 200,
        'MATERIAL': 100,
        'DESCRIPCIÓN MATERIAL': 300,
        'DESCRIPCION MATERIAL': 300,
        'FECHA INGRESO': 65,
        'CANT. IN': 30,
        'COSTO UNI': 40,
        'PVP S/IVA': 50,
        'COSTO TOTAL': 50,
        'PVP TOTAL': 50,
        'VALOR A PAGAR': 65,
        'UTILIDAD': 65,
        '% UTILIDAD': 50,
      };
      const DEFAULT_COLUMN_WIDTH = 50;
      
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
        
        const getWrappedLines = (text: string, font: PDFFont, size: number, maxWidth: number): string[] => {
            if (!text || text.trim() === '') return [''];
            
            const lines: string[] = [];
            const textBlocks = text.split('\n');

            for(const block of textBlocks) {
                const words = block.split(' ');
                let currentLine = '';
                for (const word of words) {
                    const testLine = currentLine ? `${currentLine} ${word}` : word;
                    if (font.widthOfTextAtSize(testLine, size) < maxWidth) {
                        currentLine = testLine;
                    } else {
                        if (currentLine !== '') lines.push(currentLine);
                        
                        let tempWord = word;
                        while (font.widthOfTextAtSize(tempWord, size) > maxWidth) {
                            let i = tempWord.length - 1;
                            while (i > 0 && font.widthOfTextAtSize(tempWord.substring(0, i), size) > maxWidth) {
                                i--;
                            }
                            if (i > 0) {
                                lines.push(tempWord.substring(0, i));
                                tempWord = tempWord.substring(i);
                            } else {
                                lines.push(tempWord); 
                                tempWord = '';
                                break;
                            }
                        }
                        currentLine = tempWord;
                    }
                }
                if (currentLine !== '') lines.push(currentLine);
            }

            return lines.length > 0 ? lines : [''];
        };

        let currentY = startY;
        const headerLineHeight = 12;
        const rowLineHeight = 11;
        const headerSize = 7;
        const rowSize = 5.8;
        const availableWidth = width - 2 * pageLayout.margin;
        
        const columnsToSum = ['Cant. In', 'Costo Uni', 'PVP S/IVA', 'Costo total', 'PVP Total', 'Valor a Pagar', 'Utilidad'];
        const columnsToAverage = ['% Utilidad'];
        const valueColumns = ['Cant. In','Costo Uni','PVP S/IVA','% Utilidad', 'Costo total', 'PVP Total', 'Valor a Pagar', 'Utilidad'];
        
        const upperHeaders = headers.map(h => String(h || '').trim().replace(/\.?$/, '').toUpperCase());
        const sumIndices = columnsToSum.map(colName => upperHeaders.indexOf(colName.toUpperCase().replace(/\.?$/, ''))).filter(i => i !== -1);
        const avgIndices = columnsToAverage.map(colName => upperHeaders.indexOf(colName.toUpperCase().replace(/\.?$/, ''))).filter(i => i !== -1);

        const totals = new Array(headers.length).fill(0);
        const avgTotals = new Array(headers.length).fill(0);
        const avgCounts = new Array(headers.length).fill(0);
        
        const centerAlignedIndices: number[] = [];
        headers.forEach((h, i) => {
            if (valueColumns.map(sc => sc.toUpperCase()).includes(String(h || '').trim().replace(/\.?$/, '').toUpperCase())) {
                centerAlignedIndices.push(i);
            }
        });
        
        const columnWidths = headers.map(h => 
          COLUMN_WIDTH_CONFIG[String(h || '').trim().replace(/\.?$/, '').toUpperCase()] || DEFAULT_COLUMN_WIDTH
        );
        const tableWidth = columnWidths.reduce((a, b) => a + b, 0);
        const scale = availableWidth / tableWidth;
        const scaledWidths = columnWidths.map(w => w * scale);

        // Draw Header
        const headerTopY = currentY;
        page.drawRectangle({ // Draw outer box
            x: pageLayout.margin,
            y: currentY - headerLineHeight,
            width: availableWidth,
            height: headerLineHeight,
            borderColor: rgb(0,0,0),
            borderWidth: 1.5,
            color: rgb(0.85, 0.95, 1),
        })
        let currentX = pageLayout.margin;
        headers.forEach((header, i) => {
          const headerText = String(header || '');
          const normalizedHeaderText = headerText.trim().replace(/\.?$/, '').toUpperCase();
          let currentHeaderSize = headerSize;
          
          const smallHeaders = ['CANT. IN', 'COSTO UNI', 'PVP S/IVA', '% UTILIDAD', 'COSTO TOTAL', 'PVP TOTAL', 'PROVEEDOR'];
          if (smallHeaders.includes(normalizedHeaderText)) {
            currentHeaderSize = 6;
          }
          if(normalizedHeaderText === 'FECHA INGRESO') {
            currentHeaderSize = 5.5;
          }

          const textY = currentY - (headerLineHeight / 2) - (currentHeaderSize / 2) + 2;
          page.drawText(headerText, { x: currentX + 3, y: textY , font: helveticaBoldFont, size: currentHeaderSize, color: rgb(0,0,0) });
          currentX += scaledWidths[i];

          if (i < headers.length - 1) { // Draw vertical dividers
            page.drawLine({
                start: {x: currentX, y: headerTopY},
                end: {x: currentX, y: headerTopY - headerLineHeight},
                thickness: 0.5,
                color: rgb(0,0,0)
            })
          }
        });
        currentY -= headerLineHeight;
        
        // Draw Rows
        for (const row of rows) {
            const rowCellLines: string[][] = row.map((cell, i) => {
                let cellValue;
                if (cell instanceof Date) {
                    const day = String(cell.getDate()).padStart(2, '0');
                    const month = String(cell.getMonth() + 1).padStart(2, '0');
                    const year = cell.getFullYear();
                    cellValue = `${day}/${month}/${year}`;
                } else {
                    cellValue = String(cell === null || cell === undefined ? '' : cell);
                }
                return getWrappedLines(cellValue, helveticaFont, rowSize, scaledWidths[i] - 6);
            });
            const maxLines = Math.max(1, ...rowCellLines.map(lines => lines.length));
            const dynamicRowHeight = maxLines * rowLineHeight;
            
            if (currentY - dynamicRowHeight < pageLayout.margin) {
                page = pdfDoc.addPage(pageLayout.size);
                drawPageHeader(page);
                currentY = height - pageLayout.margin - 15;

                const newHeaderTopY = currentY;
                page.drawRectangle({ // Draw outer box
                    x: pageLayout.margin,
                    y: currentY - headerLineHeight,
                    width: availableWidth,
                    height: headerLineHeight,
                    borderColor: rgb(0,0,0),
                    borderWidth: 1.5,
                    color: rgb(0.85, 0.95, 1),
                });
                let newX = pageLayout.margin;
                headers.forEach((header, i) => {
                  const headerText = String(header || '');
                  const normalizedHeaderText = headerText.trim().replace(/\.?$/, '').toUpperCase();
                  let currentHeaderSize = headerSize;
                  const smallHeaders = ['CANT. IN', 'COSTO UNI', 'PVP S/IVA', '% UTILIDAD', 'COSTO TOTAL', 'PVP TOTAL', 'PROVEEDOR'];
                  if (smallHeaders.includes(normalizedHeaderText)) {
                    currentHeaderSize = 6;
                  }
                  if(normalizedHeaderText === 'FECHA INGRESO') {
                    currentHeaderSize = 5.5;
                  }
                  const textY = currentY - (headerLineHeight / 2) - (currentHeaderSize / 2) + 2;
                  page.drawText(String(header || ''), { x: newX + 3, y: textY, font: helveticaBoldFont, size: currentHeaderSize, color: rgb(0,0,0) });
                  newX += scaledWidths[i];

                  if (i < headers.length - 1) { // Draw vertical dividers
                    page.drawLine({
                        start: {x: newX, y: newHeaderTopY},
                        end: {x: newX, y: newHeaderTopY - headerLineHeight},
                        thickness: 0.5,
                        color: rgb(0,0,0)
                    })
                  }
                });
                currentY -= headerLineHeight;
            }
            const rowTopY = currentY;
            
            let cellX = pageLayout.margin;
            rowCellLines.forEach((lines, i) => {
                const isCenterAligned = centerAlignedIndices.includes(i);
                
                lines.forEach((line, lineIndex) => {
                    const textWidth = helveticaFont.widthOfTextAtSize(line, rowSize);
                    let xPos = cellX + 3;
                    if (isCenterAligned) {
                        xPos = cellX + (scaledWidths[i] - textWidth) / 2;
                    }
                    const yPos = rowTopY - (dynamicRowHeight / 2) - ((lines.length-1) * (rowSize+1)/2) + ((lines.length-1 - lineIndex) * (rowSize+1));

                    page.drawText(line, { x: xPos, y: yPos -1 , font: helveticaFont, size: rowSize, color: rgb(0.2, 0.2, 0.2) });
                });
                cellX += scaledWidths[i];
            });

            row.forEach((cell, i) => {
                const cellValue = String(cell === null || cell === undefined ? '' : cell);
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
            });

            currentY -= dynamicRowHeight;
            
            const rowBottomY = currentY;
            page.drawLine({
              start: { x: pageLayout.margin, y: rowBottomY },
              end: { x: width - pageLayout.margin, y: rowBottomY },
              color: rgb(0, 0, 0),
              thickness: 0.5
            });
            
            let vLineX = pageLayout.margin;
            for(let i=0; i <= scaledWidths.length; i++) {
                page.drawLine({ start: {x: vLineX, y: rowTopY}, end: {x: vLineX, y: rowBottomY}, color: rgb(0, 0, 0), thickness: 0.5});
                if (i < scaledWidths.length) vLineX += scaledWidths[i];
            }
        }

        const summaryRowY = currentY;
        if (summaryRowY - headerLineHeight < pageLayout.margin) {
          page = pdfDoc.addPage(pageLayout.size);
          drawPageHeader(page);
          currentY = height - pageLayout.margin - 15;
        }


        // Draw summary row
        const summaryTopY = currentY;
        page.drawLine({ start: { x: pageLayout.margin, y: summaryTopY }, end: { x: width - pageLayout.margin, y: summaryTopY }, thickness: 1.5, color: rgb(0, 0, 0) });
        page.drawRectangle({ x: pageLayout.margin, y: currentY - headerLineHeight, width: availableWidth, height: headerLineHeight, color: rgb(1, 1, 0.8) });
        let summaryX = pageLayout.margin;
        page.drawText('*', { x: summaryX + 3, y: currentY - (headerLineHeight/2) - (rowSize/2) + 2, font: helveticaBoldFont, size: rowSize });
        
        headers.forEach((h, i) => {
            let textToDraw = '';
            let specialSize = rowSize;
            
            if (sumIndices.includes(i)) {
                 textToDraw = totals[i].toFixed(2);
                 const normalizedHeader = String(h || '').trim().replace(/\.?$/, '').toUpperCase();
                 if (normalizedHeader === 'COSTO TOTAL') {
                    specialSize = rowSize + 2;
                 }
            } else if (avgIndices.includes(i) && avgCounts[i] > 0) {
                 const average = avgTotals[i] / avgCounts[i];
                 textToDraw = average.toFixed(2);
            }

            if (textToDraw) {
                const isCenterAligned = centerAlignedIndices.includes(i);
                const textWidth = helveticaBoldFont.widthOfTextAtSize(textToDraw, specialSize);
                let xPos = summaryX + 3;
                if (isCenterAligned) {
                    xPos = summaryX + (scaledWidths[i] - textWidth) / 2;
                }
                const textY = currentY - (headerLineHeight/2) - (specialSize/2) + 2;
                page.drawText(textToDraw, { x: xPos, y: textY, font: helveticaBoldFont, size: specialSize });
            }
            summaryX += scaledWidths[i];
        });
        currentY -= headerLineHeight;
        
        const summaryBottomY = currentY;
        page.drawLine({ start: { x: pageLayout.margin, y: summaryBottomY }, end: { x: width - pageLayout.margin, y: summaryBottomY }, thickness: 0.5, color: rgb(0, 0, 0) });
        let vLineX = pageLayout.margin;
        for(let i=0; i <= scaledWidths.length; i++) {
            page.drawLine({ start: {x: vLineX, y: summaryTopY}, end: {x: vLineX, y: summaryBottomY}, color: rgb(0, 0, 0), thickness: 0.5});
            if (i < scaledWidths.length) vLineX += scaledWidths[i];
        }


        return { finalY: currentY, finalPage: page };
      };

      let pagesAdded = 0;

      for (const belnr of sortedBelnrs) {
        const ebeln = belnrToEbelnMap.get(belnr);
        if (!ebeln) continue;

        if (processedEbelns.has(ebeln)) {
          continue;
        }

        const poData = groupedByPurchaseOrder[ebeln];
        if (!poData || poData.length === 0) continue;

        processedEbelns.add(ebeln);

        if (pagesAdded > 0) {
          currentPage = pdfDoc.addPage(pageLayout.size);
          y = height - pageLayout.margin - 15;
        }
        drawPageHeader(currentPage);
        
        const { finalPage } = drawTable(currentPage, comprasHeaders, poData, y, ebeln, belnr);
        currentPage = finalPage;
        pagesAdded++;
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
          placeholder="Subir Reporte de Utilidad"
          accept=".xlsx, .xls"
          icon={<FileSpreadsheet className="h-12 w-12" />}
        />
        <FileInput
          file={documentosFile}
          onFileChange={setDocumentosFile}
          placeholder="Subir Reporte Tabla EKBE"
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
            <p className="text-muted-foreground">Tu reporte en PDF se ha descargado automáticamente.</p>
            <div className="flex justify-center gap-4">
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
                  <strong>Subir Reporte de Utilidad:</strong> Carga tu reporte principal de utilidad en formato Excel (.xlsx o .xls).
                </li>
                 <li>
                  <strong>Subir Reporte Tabla EKBE:</strong> Carga tu reporte de la tabla EKBE. Puede ser un archivo Excel (.xlsx, .xls) o el archivo exportado directamente desde SAP (generalmente con extensión .xls pero formato HTML). La aplicación intentará leer ambos formatos.
                </li>
                <li>
                  <strong>Enlace de Datos:</strong> La aplicación asocia los números de documento (BELNR) del segundo archivo con sus órdenes de compra (EBELN / Ord. de Compra) correspondientes en el primer archivo.
                </li>
                <li>
                  <strong>Generación de PDF:</strong> Se crea un único documento PDF. El reporte está organizado por número de documento, y cada sección contiene la tabla de artículos de la orden de compra asociada.
                </li>
                <li>
                  <strong>Descarga:</strong> Finalmente, tu PDF consolidado se descarga automáticamente.
                </li>
              </ol>
            </AccordionContent>
          </AccordionItem>
          <AccordionItem value="item-2">
            <AccordionTrigger>
              <div className="flex items-center gap-2 font-semibold">
                <HelpCircle className="h-5 w-5" />
                Instrucciones para obtener los archivos de SAP
              </div>
            </AccordionTrigger>
            <AccordionContent>
              <div className="space-y-4 text-sm text-muted-foreground">
                <div>
                  <h4 className="font-semibold text-foreground mb-2">Paso Previo: Obtener Números de Documento (BELNR)</h4>
                  <ol className="list-decimal space-y-2 pl-6">
                    <li>
                      <strong>Transacción <code>MIR5</code></strong>:
                      <ul className="list-disc pl-5 mt-1 space-y-1">
                        <li>Ingresa, filtra y descarga las facturas según el rango de fechas, usuario y sociedad.</li>
                        <li>Copia o exporta los números de documento (<code>BELNR</code>), los necesitarás para los siguientes pasos.</li>
                      </ul>
                    </li>
                  </ol>
                </div>

                <div>
                  <h4 className="font-semibold text-foreground mb-2">Archivo 1: Reporte de Utilidad (desde ZREP PEDIDOS)</h4>
                  <ol className="list-decimal space-y-2 pl-6">
                    <li>
                      <strong>Transacción <code>ZREP PEDIDOS</code></strong>:
                      <ul className="list-disc pl-5 mt-1 space-y-1">
                          <li>En el campo de selección de "Facturas", pega todos los números de documento que obtuviste.</li>
                          <li>Ejecuta el reporte (<strong>F8</strong>).</li>
                          <li>Exportar como "Hoja de cálculo" y luego guardar el archivo en formato Excel.</li>
                      </ul>
                    </li>
                  </ol>
                </div>

                <div>
                  <h4 className="font-semibold text-foreground mb-2">Archivo 2: Reporte Tabla EKBE (desde SE16)</h4>
                   <ol className="list-decimal space-y-2 pl-6">
                     <li>
                       <strong>Transacción <code>SE16</code></strong>:
                       <ul className="list-disc pl-5 mt-1 space-y-1">
                         <li>Ingresa a la transacción, escribe la tabla <strong><code>EKBE</code></strong> y presiona <strong>Enter</strong>.</li>
                         <li>Carga la variante: Menú <em>Pasar a &gt; Variantes &gt; Traer...</em> y selecciona <strong><code>REVOC</code></strong>.</li>
                         <li>En el campo <code>BELNR</code>, usa la selección múltiple para pegar todos los números de documento de la <code>MIR5</code>.</li>
                         <li>Ejecuta la selección (<strong>F8</strong>).</li>
                         <li>Asegúrate de que las columnas <code>BELNR</code> y <code>EBELN</code> estén visibles.</li>
                         <li>Exporta la lista (por ejemplo, desde el menú <em>Sistema &gt; Lista &gt; Grabar &gt; Fichero local</em>) eligiendo la opción "Hoja de cálculo".</li>
                         <li>Guarda el archivo resultante. Este será el archivo que subirás como segundo reporte.</li>
                       </ul>
                     </li>
                   </ol>
                </div>
              </div>
            </AccordionContent>
          </AccordionItem>
        </Accordion>
      </div>
    </div>
  );
}
