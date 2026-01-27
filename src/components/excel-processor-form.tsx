"use client";

import * as React from "react";
import * as XLSX from "xlsx";
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
  const [excelFile, setExcelFile] = React.useState<File | null>(null);
  const [status, setStatus] = React.useState<Status>("idle");
  const [progressMessage, setProgressMessage] = React.useState("");
  const [downloadUrl, setDownloadUrl] = React.useState<string | null>(null);
  const { toast } = useToast();

  const resetState = () => {
    setExcelFile(null);
    setStatus("idle");
    setProgressMessage("");
    setDownloadUrl(null);
  };
  
  const handleProcess = async () => {
    if (!excelFile) {
      toast({
        variant: "destructive",
        title: "Archivo Faltante",
        description: "Por favor, sube un archivo de Excel.",
      });
      return;
    }

    setStatus("processing");

    try {
      setProgressMessage("Leyendo archivo de Excel...");
      const excelBuffer = await excelFile.arrayBuffer();
      const workbook = XLSX.read(excelBuffer, { type: "buffer" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const headerRowIndex = data.findIndex(row => Array.isArray(row) && row.includes("Ord. de Compra"));
      if (headerRowIndex === -1) {
        throw new Error("No se encontró la fila de encabezado con 'Ord. de Compra'.");
      }
      
      const headers = data[headerRowIndex];
      const purchaseOrderColIndex = headers.indexOf("Ord. de Compra");
      if (purchaseOrderColIndex === -1) {
        // This case should be covered by the headerRowIndex check, but it's good practice.
        throw new Error("No se encontró la columna 'Ord. de Compra' en el archivo.");
      }

      setProgressMessage("Agrupando por orden de compra...");
      const dataRows = data.slice(headerRowIndex + 1);
      const groupedByPurchaseOrder = dataRows.reduce((acc, row) => {
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

      if (Object.keys(groupedByPurchaseOrder).length === 0) {
        setStatus("error");
        toast({
            variant: "destructive",
            title: "No se Encontraron Datos",
            description: "No se encontraron órdenes de compra para procesar en el archivo Excel.",
        });
        return;
      }

      setProgressMessage("Generando nuevo archivo Excel...");
      const newWorkbook = XLSX.utils.book_new();
      
      for (const purchaseOrder in groupedByPurchaseOrder) {
        const rowsForPO = groupedByPurchaseOrder[purchaseOrder];
        const sheetData = [headers, ...rowsForPO];
        const newSheet = XLSX.utils.aoa_to_sheet(sheetData);
        
        const sanitizedSheetName = purchaseOrder.replace(/[\\/?*[\]]/g, "").substring(0, 31);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, sanitizedSheetName);
      }

      const newExcelBuffer = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });

      const blob = new Blob([newExcelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
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
          error instanceof Error ? error.message : "No se pudo procesar el archivo. Por favor, revisa que el formato sea correcto.",
      });
    }
  };

  const isIdle = status === 'idle';
  const isProcessing = status === 'processing';
  const isSuccess = status === 'success';
  const isError = status === 'error';

  return (
    <div className="space-y-6">
      <div className="flex justify-center">
        <FileInput
          file={excelFile}
          onFileChange={setExcelFile}
          placeholder="Subir Reporte de Compras"
          accept=".xlsx, .xls"
          icon={<FileSpreadsheet className="h-12 w-12" />}
        />
      </div>

      <div className="flex justify-center">
        {isIdle && (
          <Button
            size="lg"
            onClick={handleProcess}
            disabled={!excelFile}
          >
            Procesar Reporte <ArrowRight className="ml-2 h-4 w-4" />
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
            <p className="text-muted-foreground">Tu nuevo reporte de Excel está listo para descargar.</p>
            <div className="flex justify-center gap-4">
               <Button size="lg" asChild>
                <a href={downloadUrl!} download="Reporte_Agrupado.xlsx">
                  <Download className="mr-2 h-4 w-4" />
                  Descargar Reporte
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
             <p className="text-muted-foreground">No pudimos procesar tu reporte. Por favor, inténtalo de nuevo.</p>
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
                  <strong>Subir Archivo:</strong> Carga tu reporte de compras en formato Excel (.xlsx o .xls).
                </li>
                <li>
                  <strong>Agrupación Automática:</strong> La aplicación lee el archivo, busca la columna "Ord. de Compra" y agrupa todas las filas según el número de orden.
                </li>
                <li>
                  <strong>Generación del Nuevo Reporte:</strong> Se crea un nuevo archivo de Excel. Cada "Ord. de Compra" única se convierte en una hoja de cálculo separada dentro de este nuevo archivo.
                </li>
                <li>
                  <strong>Descarga:</strong> Finalmente, se te proporciona un enlace para descargar el nuevo archivo de Excel con todos los datos organizados por hojas.
                </li>
              </ol>
            </AccordionContent>
          </AccordionItem>
        </Accordion>
      </div>
    </div>
  );
}
