'use server';

/**
 * @fileOverview Groups PDF pages by the EBELN (purchase order number) found on them.
 *
 * - groupPagesByEbeln - A function that handles the page grouping process.
 * - GroupPagesByEbelnInput - The input type for the groupPagesByEbeln function.
 * - GroupPagesByEbelnOutput - The return type for the groupPagesByEbeln function.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const GroupPagesByEbelnInputSchema = z.object({
  pdfPages: z.array(
    z.object({
      pageNumber: z.number().describe('The page number of the PDF.'),
      pageText: z.string().describe('The text content of the PDF page.'),
    })
  ).describe('An array of PDF pages with their text content and page numbers.'),
});
export type GroupPagesByEbelnInput = z.infer<typeof GroupPagesByEbelnInputSchema>;

const GroupPagesByEbelnOutputSchema = z.array(
  z.object({
    ebeln: z.string().describe("The exact EBELN value (número de pedido) as it appears in the document."),
    pageNumbers: z.array(z.number()).describe("An array of page numbers belonging to this EBELN."),
  })
).describe("An array of objects where each object represents a document, containing the EBELN and all associated page numbers.");
export type GroupPagesByEbelnOutput = z.infer<typeof GroupPagesByEbelnOutputSchema>;

export async function groupPagesByEbeln(
  input: GroupPagesByEbelnInput
): Promise<GroupPagesByEbelnOutput> {
  return groupPagesByEbelnFlow(input);
}

const groupPagesPrompt = ai.definePrompt({
  name: 'groupPagesPrompt',
  input: {schema: GroupPagesByEbelnInputSchema},
  output: {schema: GroupPagesByEbelnOutputSchema},
  prompt: `You are an expert data extraction specialist. Your task is to group PDF pages by the Purchase Order Number (in Spanish: 'número de pedido' or 'EBELN') found on them.

  - The 'número de pedido' is the key identifier. You must extract it **exactly** as it appears in the text, including any leading zeros. Do not alter or normalize it.
  - A document for a single 'número de pedido' can span multiple pages.
  - Often, the 'número de pedido' is only on the first page of a multi-page document. All subsequent pages belong to that same 'número de pedido' until a new one is found on a later page.

  Analyze the text from all pages provided. Return an array of objects. Each object must contain:
  1.  'ebeln': The exact 'número de pedido' string.
  2.  'pageNumbers': An array of all page numbers (as integers) that belong to that 'número de pedido'.

  Example of the expected output format:
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

  Here is the text content from the PDF pages:
  {{#each pdfPages}}
  Page Number: {{{pageNumber}}}
  Page Text: {{{pageText}}}
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
    const {output} = await groupPagesPrompt(input);
    return output!;
  }
);
