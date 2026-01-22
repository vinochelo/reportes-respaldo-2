'use server';

/**
 * @fileOverview Extrae el número de pedido (EBELN) de una sola página de texto de un PDF.
 *
 * - extractEbelnFromPage - Una función que maneja el proceso de extracción de EBELN.
 * - ExtractEbelnFromPageInput - El tipo de entrada para la función extractEbelnFromPage.
 * - ExtractEbelnFromPageOutput - El tipo de retorno para la función extractEbelnFromPage.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const ExtractEbelnFromPageInputSchema = z.object({
  pageText: z.string().describe('El contenido de texto de una página de PDF.'),
});
export type ExtractEbelnFromPageInput = z.infer<typeof ExtractEbelnFromPageInputSchema>;

const ExtractEbelnFromPageOutputSchema = z.object({
    ebeln: z.string().nullable().describe("El valor de EBELN (número de pedido) encontrado en la página, o null si no se encuentra."),
});
export type ExtractEbelnFromPageOutput = z.infer<typeof ExtractEbelnFromPageOutputSchema>;


export async function extractEbelnFromPage(
  input: ExtractEbelnFromPageInput
): Promise<ExtractEbelnFromPageOutput> {
  return extractEbelnFromPageFlow(input);
}


const extractEbelnPrompt = ai.definePrompt({
  name: 'extractEbelnFromPagePrompt',
  model: 'gemini-1.0-pro',
  input: {schema: ExtractEbelnFromPageInputSchema},
  output: {schema: ExtractEbelnFromPageOutputSchema},
  prompt: `Tu tarea es encontrar y extraer el número de pedido 'EBELN' del siguiente texto.
El valor de EBELN es un número.
Analiza el texto y encuentra el valor asociado con la etiqueta 'EBELN'.
Si encuentras un valor para EBELN, devuélvelo como un string. Si no encuentras ningún valor de EBELN, devuelve null.
Devuelve SÓLO el JSON. No incluyas texto antes o después.

Texto de la página:
{{{pageText}}}
`,
});

const extractEbelnFromPageFlow = ai.defineFlow(
  {
    name: 'extractEbelnFromPageFlow',
    inputSchema: ExtractEbelnFromPageInputSchema,
    outputSchema: ExtractEbelnFromPageOutputSchema,
  },
  async input => {
    const {output} = await extractEbelnPrompt(input);
    return output!;
  }
);
