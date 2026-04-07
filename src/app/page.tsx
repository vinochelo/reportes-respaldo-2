"use client";

import * as React from "react";
import { ExcelProcessorForm } from "@/components/excel-processor-form";
import { FileText } from "lucide-react";
import { Card } from "@/components/ui/card";

export default function Home() {
  const [year, setYear] = React.useState<number | null>(null);

  React.useEffect(() => {
    setYear(new Date().getFullYear());
  }, []);

  return (
    <div className="min-h-screen bg-[#fafafa] relative overflow-hidden font-sans selection:bg-blue-100">
      {/* Grid Background */}
      <div 
        className="absolute inset-0 pointer-events-none z-0" 
        style={{
          backgroundImage: `
            linear-gradient(to right, rgba(0,0,0,0.03) 1px, transparent 1px),
            linear-gradient(to bottom, rgba(0,0,0,0.03) 1px, transparent 1px)
          `,
          backgroundSize: '40px 40px',
          maskImage: 'linear-gradient(to bottom, white 40%, transparent)',
          WebkitMaskImage: 'linear-gradient(to bottom, white 40%, transparent)'
        }}
      />

      <div className="relative z-10 p-6 md:p-12 max-w-5xl mx-auto flex flex-col items-center">
        
        {/* Header Section */}
        <div className="text-center space-y-4 mb-10 pt-4 md:pt-8">
          <div className="mx-auto w-16 h-16 bg-blue-50 rounded-2xl flex items-center justify-center mb-6 shadow-sm border border-blue-100">
            <FileText className="h-8 w-8 text-blue-500" />
          </div>
          
          <h1 className="text-4xl md:text-5xl font-bold tracking-tight text-gray-900">
            Generador de <span className="text-blue-600">Reportes</span>
          </h1>
          
          <p className="max-w-xl mx-auto text-base text-gray-500 mt-4 leading-relaxed">
            Sube el reporte de <span className="font-medium text-gray-700">Utilidad</span> y la tabla <span className="font-medium text-gray-700">EKBE</span> en Excel y obtén los reportes individuales listos para imprimir en PDF de forma automática.
          </p>
        </div>

        {/* Main Card */}
        <div className="w-full max-w-4xl">
          <Card className="bg-white/80 backdrop-blur-xl border-gray-100 shadow-xl shadow-gray-200/50 p-6 md:p-12 rounded-3xl">
            <ExcelProcessorForm />
          </Card>
        </div>

        {/* Footer */}
        {year && (
          <footer className="mt-16 text-center text-sm text-gray-400">
            © {year} Rocku. Todos los derechos reservados.
          </footer>
        )}
      </div>
    </div>
  );
}
