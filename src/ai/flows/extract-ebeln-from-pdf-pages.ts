'use server';
/**
 * @fileOverview Extracts an EBELN (purchase order number) from a text snippet.
 *
 * - extractEbeln - A function that handles the EBELN extraction process.
 * - ExtractEbelnInput - The input type for the extractEbeln function.
 * - ExtractEbelnOutput - The return type for the extractEbeln function.
 */

import {ai} from '@/ai/genkit';
import {googleAI} from '@genkit-ai/google-genai';
import {z} from 'genkit';

const ExtractEbelnInputSchema = z.object({
  pageText: z.string().describe('A snippet of text from the top of a PDF page.'),
});
export type ExtractEbelnInput = z.infer<typeof ExtractEbelnInputSchema>;

const ExtractEbelnOutputSchema = z.object({
  ebeln: z.string().nullable().describe('The identified EBELN (purchase order number), or null if not found.'),
});
export type ExtractEbelnOutput = z.infer<typeof ExtractEbelnOutputSchema>;


const extractEbelnPrompt = ai.definePrompt({
  name: 'extractEbelnPrompt',
  model: googleAI.model('gemini-flash'),
  input: {schema: ExtractEbelnInputSchema},
  output: {schema: ExtractEbelnOutputSchema},
  prompt: `You are an expert at finding purchase order numbers in document text. The purchase order number is often labeled as "Pedido", "EBELN", "PO No.", or "Orden de Compra". It can be a purely numeric value (like 4500123456) or alphanumeric. Find the purchase order number in the following text and return it. If no number is found, return null.

Text:
\'\'\'
{{{pageText}}}
\'\'\'`,
});

const extractEbelnFlow = ai.defineFlow(
  {
    name: 'extractEbelnFlow',
    inputSchema: ExtractEbelnInputSchema,
    outputSchema: ExtractEbelnOutputSchema,
  },
  async input => {
    const {output} = await extractEbelnPrompt(input);
    return output!;
  }
);

export async function extractEbeln(input: ExtractEbelnInput): Promise<ExtractEbelnOutput> {
  return extractEbelnFlow(input);
}
