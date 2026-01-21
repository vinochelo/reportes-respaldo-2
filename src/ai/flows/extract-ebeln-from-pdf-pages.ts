
'use server';

/**
 * @fileOverview Agrupa las páginas de un PDF según el EBELN (número de pedido) que se encuentra en ellas.
 *
 * - groupPagesByEbeln - Una función que maneja el proceso de agrupación de páginas.
 * - GroupPagesByEbelnInput - El tipo de entrada para la función groupPagesByEbeln.
 * - GroupPagesByEbelnOutput - El tipo de retorno para la función groupPagesByEbeln.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const GroupPagesByEbelnInputSchema = z.object({
  pdfPages: z.array(
    z.object({
      pageNumber: z.number().describe('El número de página del PDF.'),
      pageText: z.string().describe('El contenido de texto de la página del PDF.'),
    })
  ).describe('Un arreglo de páginas de PDF con su contenido de texto y números de página.'),
});
export type GroupPagesByEbelnInput = z.infer<typeof GroupPagesByEbelnInputSchema>;

const GroupPagesByEbelnOutputSchema = z.array(
  z.object({
    ebeln: z.string().describe("El valor exacto de EBELN (número de pedido) tal como aparece en el documento."),
    pageNumbers: z.array(z.number()).describe("Un arreglo de números de página que pertenecen a este EBELN."),
  })
).describe("Un arreglo de objetos donde cada objeto representa un documento, conteniendo el EBELN y todos los números de página asociados.");
export type GroupPagesByEbelnOutput = z.infer<typeof GroupPagesByEbelnOutputSchema>;

export async function groupPagesByEbeln(
  input: GroupPagesByEbelnInput
): Promise<GroupPagesByEbelnOutput> {
  return groupPagesByEbelnFlow(input);
}

const extractEbelnPrompt = ai.definePrompt({
  name: 'extractEbelnPrompt',
  model: 'gemini-1.5-pro',
  input: {schema: GroupPagesByEbelnInputSchema},
  output: {schema: GroupPagesByEbelnOutputSchema},
  prompt: `Tu tarea es agrupar las páginas de un PDF por el número de pedido 'EBELN'.

Extrae el valor exacto de 'EBELN' de cada documento.
Agrupa todos los números de página que pertenecen a cada 'EBELN'.
Devuelve un arreglo de objetos JSON con 'ebeln' (string) y 'pageNumbers' (array de números).

Páginas:
{{#each pdfPages}}
Página {{{pageNumber}}}: {{{pageText}}}
---
{{/each}}`,
});

const groupPagesByEbelnFlow = ai.defineFlow(
  {
    name: 'groupPagesByEbelnFlow',
    inputSchema: GroupPagesByEbelnInputSchema,
    outputSchema: GroupPagesByEbelnOutputSchema,
  },
  async input => {
    const {output} = await extractEbelnPrompt(input);
    return output!;
  }
);
