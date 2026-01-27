"use client";

import * as React from "react";
import { ExcelProcessorForm } from '@/components/excel-processor-form';
import { Card, CardContent } from '@/components/ui/card';

export default function Home() {
  const [year, setYear] = React.useState<number | null>(null);

  React.useEffect(() => {
    setYear(new Date().getFullYear());
  }, []);

  return (
    <div className="flex min-h-dvh flex-col items-center justify-center bg-background px-4 py-12 sm:px-6 lg:px-8">
      <div className="mx-auto w-full max-w-3xl space-y-8">
        <div className="text-center">
          <h1 className="text-5xl font-extrabold tracking-tight text-foreground sm:text-6xl">
            <span className="font-light">ROCKU</span>
          </h1>
          <p className="mt-4 text-xl text-muted-foreground">
            Generador de Reportes PDF
          </p>
        </div>
        <Card className="shadow-lg">
          <CardContent className="p-8">
            <ExcelProcessorForm />
          </CardContent>
        </Card>
        {year && (
          <footer className="text-center text-sm text-muted-foreground">
            © {year} Rocku. Todos los derechos reservados.
          </footer>
        )}
      </div>
    </div>
  );
}
