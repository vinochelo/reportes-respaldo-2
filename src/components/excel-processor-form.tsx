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
  title: string;
  subtitle: string;
  accept: string;
  icon: React.ReactNode;
}

const FileInput: React.FC<FileInputProps> = ({
  file,
  onFileChange,
  title,
  subtitle,
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
        "group relative flex flex-col items-center justify-center w-full py-10 px-6 border-2 border-dashed rounded-3xl cursor-pointer transition-all duration-300",
        file 
          ? "border-blue-500 bg-blue-50/50" 
          : "border-gray-200 hover:border-blue-500/50 hover:bg-blue-50/30 bg-white"
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
        <div className="text-center animate-in fade-in zoom-in duration-300">
          <div className="w-16 h-16 rounded-full bg-blue-100 flex items-center justify-center mx-auto mb-4">
            <div className="text-blue-600">{icon}</div>
          </div>
          <p className="mt-2 text-lg font-semibold text-gray-900">{file.name}</p>
          <p className="text-sm text-gray-500 mt-1">
            {Math.round(file.size / 1024)} KB
          </p>
          <Button
            variant="ghost"
            size="sm"
            className="mt-4 text-rose-500 hover:text-rose-600 hover:bg-rose-50 rounded-full px-4"
            onClick={handleRemoveFile}
          >
            <X className="mr-2 h-4 w-4" /> Quitar archivo
          </Button>
        </div>
      ) : (
        <div className="text-center">
          <div className="w-20 h-20 rounded-full bg-blue-50 flex items-center justify-center mx-auto mb-6 group-hover:scale-110 transition-transform duration-300">
            <UploadCloud className="h-10 w-10 text-blue-500" />
          </div>
          <h3 className="text-xl font-semibold text-gray-900 mb-2">{title}</h3>
          <p className="text-sm text-gray-500 mb-6 max-w-[280px] mx-auto">{subtitle}</p>
          <Button 
            variant="default"
            className="bg-blue-600 hover:bg-blue-700 text-white rounded-full px-8 shadow-sm shadow-blue-200"
            onClick={(e) => {
              e.stopPropagation();
              inputRef.current?.click();
            }}
          >
            {icon}
            <span className="ml-2">Seleccionar Archivo</span>
          </Button>
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
      setProgressMessage("Analizando reporte de documentos...");
      
      const findHeaders = (data: any[][]): { headerRow: number, ebelnCol: number, belnrCol: number } | null => {
          if (!data || data.length === 0) return null;
          for (let i = 0; i < Math.min(data.length, 50); i++) {
              const row = data[i];
              if (!Array.isArray(row)) continue;
              
              const normalizedRow = row.map(cell => String(cell || '').toLowerCase().replace(/[\s._-]/g, ''));
              
              const ebelnIndex = normalizedRow.findIndex(text => text.includes('ebeln') || text.includes('doccompr') || text.includes('pedido'));
              const belnrIndex = normalizedRow.findIndex(text => text.includes('belnr') || text.includes('docmat') || text.includes('nrodoc'));

              if (ebelnIndex !== -1 && belnrIndex !== -1) {
                  return { headerRow: i, ebelnCol: ebelnIndex, belnrCol: belnrIndex };
              }
          }
          return null;
      };

      const parseDocumentsFile = async (file: File): Promise<{ docData: any[][], headers: { headerRow: number; ebelnCol: number; belnrCol: number } } | null> => {
        const buffer = await file.arrayBuffer();
        let workbook;

        // The 'xlsx' library can handle multiple formats. We try binary first for .xlsx/.xls,
        // then fall back to text for formats like TSV ("Texto con tabuladores").
        try {
            setProgressMessage("Intentando leer como archivo Excel...");
            workbook = XLSX.read(buffer, { type: 'buffer' });
        } catch (e) {
            console.warn("Fallo al leer como Excel binario. Se intentará como texto.", e);
            try {
                setProgressMessage("Intentando leer como archivo de texto...");
                const textData = new TextDecoder('utf-8').decode(buffer);
                workbook = XLSX.read(textData, { type: 'string' });
            } catch (textError) {
                console.error("Fallo definitivo al procesar el archivo de documentos.", textError);
                throw new Error("No se pudo leer el archivo de documentos. Asegúrate de que no esté corrupto y sea un formato válido (Excel o Texto con tabuladores).");
            }
        }

        if (!workbook || !workbook.SheetNames.length) {
            return null;
        }

        let largestSheetData: any[][] = [];
        let maxRows = 0;
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const sheetData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            if (sheetData.length > maxRows) {
                largestSheetData = sheetData;
                maxRows = sheetData.length;
            }
        }
        
        if (largestSheetData.length > 0) {
            const headers = findHeaders(largestSheetData);
            if (headers) {
                setProgressMessage("Archivo de documentos procesado.");
                return { docData: largestSheetData, headers };
            }
        }
        
        return null;
      };
      
      const parseResult = await parseDocumentsFile(documentosFile);

      if (!parseResult) {
          throw new Error("No se encontró la fila de encabezado con ('EBELN' o 'Doc.compr.') y ('BELNR' o 'Doc.mat.') en el archivo de documentos. Por favor, asegúrate de que el archivo es correcto y no está dañado.");
      }

      const { docData, headers } = parseResult;
      const {headerRow: docHeaderRowIndex, ebelnCol: ebelnColIndex, belnrCol: belnrColIndex } = headers;

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
        size: [PageSizes.A4[1], PageSizes.A4[0]] as [number, number], // Landscape A4
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
            const textBlocks = text.split('\\n');

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
          const originalHeaderText = String(header || '');
          let headerText = originalHeaderText;
          const normalizedHeaderForCheck = originalHeaderText.trim().toUpperCase();

          if (normalizedHeaderForCheck.startsWith("ORD. DE COMPRA")) {
            headerText = 'Ord. de Com.';
          } else if (normalizedHeaderForCheck.startsWith("CANT. IN")) {
            headerText = 'Cant. .';
          } else if (normalizedHeaderForCheck.startsWith("COSTO UNI")) {
            headerText = 'Costo.';
          }
          
          const normalizedHeaderText = originalHeaderText.trim().replace(/\.?$/, '').toUpperCase();
          let currentHeaderSize = headerSize;
          
          const smallHeaders = ['ORD. DE COMPRA', 'CANT. IN', 'COSTO UNI', 'COSTO TOTAL', 'VALOR A PAGAR', 'PVP S/IVA', '% UTILIDAD', 'PVP TOTAL', 'PROVEEDOR'];
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
                  const originalHeaderText = String(header || '');
                  let headerText = originalHeaderText;
                  const normalizedHeaderForCheck = originalHeaderText.trim().toUpperCase();
                  if (normalizedHeaderForCheck.startsWith("ORD. DE COMPRA")) {
                    headerText = 'Ord. de Com.';
                  } else if (normalizedHeaderForCheck.startsWith("CANT. IN")) {
                    headerText = 'Cant. .';
                  } else if (normalizedHeaderForCheck.startsWith("COSTO UNI")) {
                    headerText = 'Costo.';
                  }

                  const normalizedHeaderText = originalHeaderText.trim().replace(/\.?$/, '').toUpperCase();
                  let currentHeaderSize = headerSize;
                  const smallHeaders = ['ORD. DE COMPRA', 'CANT. IN', 'COSTO UNI', 'COSTO TOTAL', 'VALOR A PAGAR', 'PVP S/IVA', '% UTILIDAD', 'PVP TOTAL', 'PROVEEDOR'];
                  if (smallHeaders.includes(normalizedHeaderText)) {
                    currentHeaderSize = 6;
                  }
                  if(normalizedHeaderText === 'FECHA INGRESO') {
                    currentHeaderSize = 5.5;
                  }
                  const textY = currentY - (headerLineHeight / 2) - (currentHeaderSize / 2) + 2;
                  page.drawText(headerText, { x: newX + 3, y: textY, font: helveticaBoldFont, size: currentHeaderSize, color: rgb(0,0,0) });
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
      const blob = new Blob([pdfBytes as any], { type: "application/pdf" });
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
    <div className="space-y-10">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        <FileInput
          file={comprasFile}
          onFileChange={setComprasFile}
          title="Haz clic para subir Reporte de Utilidad"
          subtitle="Formato .xlsx o .xls aceptado. Archivo desde ZREP PEDIDOS."
          accept=".xlsx, .xls"
          icon={<FileSpreadsheet className="h-5 w-5" />}
        />
        <FileInput
          file={documentosFile}
          onFileChange={setDocumentosFile}
          title="Haz clic para subir Tabla EKBE"
          subtitle="Formato .txt o .xlsx aceptado. Archivo desde SE16."
          accept=".xlsx, .xls, .txt"
          icon={<FileText className="h-5 w-5" />}
        />
      </div>

      <div className="flex justify-center">
        {isIdle && (
          <Button
            size="lg"
            className="w-full md:w-auto px-12 py-6 text-lg rounded-full bg-blue-600 hover:bg-blue-700 shadow-lg shadow-blue-200"
            onClick={handleProcess}
            disabled={!comprasFile || !documentosFile}
          >
            Generar PDF <ArrowRight className="ml-2 h-5 w-5" />
          </Button>
        )}

        {isProcessing && (
          <div className="flex flex-col items-center space-y-4">
            <Loader2 className="h-10 w-10 animate-spin text-blue-500" />
            <div className="text-lg font-medium text-gray-700">
              {progressMessage}
            </div>
          </div>
        )}

        {isSuccess && (
          <div className="text-center space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="mx-auto w-20 h-20 bg-blue-100 rounded-full flex items-center justify-center mb-2">
              <PartyPopper className="h-10 w-10 text-blue-600" />
            </div>
            <div>
              <h3 className="text-3xl font-bold text-gray-900 mb-2">¡Proceso Completado!</h3>
              <p className="text-gray-500 text-lg">Tu reporte en PDF se ha descargado automáticamente.</p>
            </div>
            <div className="flex justify-center pt-2">
              <Button size="lg" className="rounded-full px-8" variant="outline" onClick={resetState}>
                Procesar Otros Archivos
              </Button>
            </div>
          </div>
        )}

        {isError && (
          <div className="text-center space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="mx-auto w-20 h-20 bg-rose-100 rounded-full flex items-center justify-center mb-2">
               <div className="text-rose-500 font-bold text-4xl">!</div>
            </div>
            <div>
              <h3 className="text-3xl font-bold text-rose-600 mb-2">¡Uy! Algo salió mal.</h3>
              <p className="text-gray-500 text-lg">No pudimos procesar tus reportes. Por favor, inténtalo de nuevo.</p>
            </div>
             <div className="flex justify-center pt-2">
               <Button size="lg" className="rounded-full px-8 bg-gray-900 text-white hover:bg-gray-800" onClick={resetState}>
                 <RefreshCw className="mr-2 h-4 w-4" />
                 Intentar de Nuevo
               </Button>
             </div>
          </div>
        )}
      </div>

      <div className="pt-10 mb-[-1rem]">
        <div className="max-w-2xl mx-auto">
          <div className="flex items-center gap-3 mb-6">
            <div className="w-8 h-8 rounded-lg bg-blue-100 flex items-center justify-center text-blue-600 font-bold">1</div>
            <h2 className="text-2xl font-semibold text-gray-900">¿Cómo funciona?</h2>
          </div>
          
          <Accordion type="single" collapsible className="w-full bg-slate-50/50 rounded-2xl px-6 border border-slate-100">
            <AccordionItem value="item-1" className="border-b border-slate-200/50">
              <AccordionTrigger className="hover:no-underline py-4 text-base font-medium text-gray-800">
                <div className="flex items-center gap-2">
                  <HelpCircle className="h-5 w-5 text-slate-400" />
                  El proceso paso a paso
                </div>
              </AccordionTrigger>
              <AccordionContent className="text-slate-600 pb-6 leading-relaxed">
                <ol className="list-decimal space-y-3 pl-6 mt-2">
                  <li>
                    <strong>Subir Reporte de Utilidad:</strong> Carga tu reporte principal de utilidad en formato Excel (.xlsx o .xls).
                  </li>
                  <li>
                    <strong>Subir Reporte Tabla EKBE:</strong> Carga tu reporte de la tabla EKBE en formato <strong>"Texto con tabuladores"</strong> (archivo <code>.txt</code>) o Excel.
                  </li>
                  <li>
                    <strong>Enlace de Datos:</strong> El sistema enlaza automáticamente los números de documento (BELNR) con sus órdenes correspondientes.
                  </li>
                  <li>
                    <strong>Generación y Descarga:</strong> Se crea un PDF formateado y organizado por orden de compra que se descargará de inmediato.
                  </li>
                </ol>
              </AccordionContent>
            </AccordionItem>
            
            <AccordionItem value="item-2" className="border-none">
              <AccordionTrigger className="hover:no-underline py-4 text-base font-medium text-gray-800">
                <div className="flex items-center gap-2">
                  <HelpCircle className="h-5 w-5 text-slate-400" />
                  Instrucciones detalladas de SAP
                </div>
              </AccordionTrigger>
              <AccordionContent className="text-slate-600 pb-6 leading-relaxed">
                <div className="space-y-6 mt-2">
                  <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-sm">
                    <h4 className="font-semibold text-gray-900 mb-2 flex items-center gap-2">
                      <span className="w-6 h-6 rounded-md bg-slate-100 flex items-center justify-center text-xs">P</span>
                      Paso Previo: Obtener BELNR
                    </h4>
                    <p className="text-sm mb-2">Transacción <code>MIR5</code>. Filtra por fechas y usuario, luego copia los números de documento (BELNR).</p>
                  </div>

                  <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-sm">
                    <h4 className="font-semibold text-gray-900 mb-2 flex items-center gap-2">
                      <span className="w-6 h-6 rounded-md bg-slate-100 flex items-center justify-center text-xs">1</span>
                      Reporte de Utilidad
                    </h4>
                    <p className="text-sm">Transacción <code>ZREP PEDIDOS</code>. Pega los documentos en 'Facturas', ejecuta (F8), y exporta a Excel.</p>
                  </div>

                  <div className="bg-white p-4 rounded-xl border border-slate-100 shadow-sm">
                    <h4 className="font-semibold text-gray-900 mb-2 flex items-center gap-2">
                      <span className="w-6 h-6 rounded-md bg-slate-100 flex items-center justify-center text-xs">2</span>
                      Tabla EKBE
                    </h4>
                    <p className="text-sm">Transacción <code>SE16</code>, tabla <code>EKBE</code>. Variante <code>REVOC</code>. Pega los documentos, ejecuta (F8), y exporta como 'Texto con tabuladores' (.txt).</p>
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>
          </Accordion>
        </div>
      </div>
    </div>
  );
}
