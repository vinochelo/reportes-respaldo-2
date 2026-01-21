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
  model: 'gemini-2.5-flash-lite',
  input: {schema: GroupPagesByEbelnInputSchema},
  output: {schema: GroupPagesByEbelnOutputSchema},
  prompt: `Eres un especialista experto en extracción de datos. Tu tarea es agrupar las páginas de un PDF por el Número de Pedido ('número de pedido' o 'EBELN') que se encuentra en ellas.

- El 'número de pedido' es el identificador clave. Debes extraerlo **exactamente** como aparece en el texto, incluyendo los ceros iniciales. No lo alteres ni lo normalices.
- Un documento para un solo 'número de pedido' puede abarcar varias páginas.
- A menudo, el 'número de pedido' solo se encuentra en la primera página de un documento de varias páginas. Todas las páginas siguientes pertenecen a ese mismo 'número de pedido' hasta que se encuentre uno nuevo en una página posterior.

Analiza el texto de todas las páginas proporcionadas. Devuelve un arreglo de objetos. Cada objeto debe contener:
1. 'ebeln': El string exacto del 'número de pedido'.
2. 'pageNumbers': Un arreglo de todos los números de página (como enteros) que pertenecen a ese 'número de pedido'.

Ejemplo del formato de salida esperado:
[
  {
    "ebeln": "4500018595",
    "pageNumbers": [1, 2, 3]
  },
  {
    "ebeln": "004500018596",
    "pageNumbers": [4]
  }
]

Aquí está el contenido de texto de las páginas del PDF:
{{#each pdfPages}}
Número de Página: {{{pageNumber}}}
Texto de la Página: {{{pageText}}}
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
