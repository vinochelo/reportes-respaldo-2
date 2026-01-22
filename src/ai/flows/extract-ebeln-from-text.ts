'use server';
/**
 * @fileOverview Extracts a purchase order number (EBELN) from a block of text using AI.
 *
 * - extractEbeln - A function that calls the AI model to find the EBELN.
 */
import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const EbelnInputSchema = z.string();
const EbelnOutputSchema = z.string().describe("The purchase order number, or an empty string if not found.");

export async function extractEbeln(input: string): Promise<string> {
  if (!input || input.trim() === '') {
    return '';
  }
  return extractEbelnFlow(input);
}

const extractEbelnFlow = ai.defineFlow(
  {
    name: 'extractEbelnFlow',
    inputSchema: EbelnInputSchema,
    outputSchema: EbelnOutputSchema,
  },
  async input => {
    const modelToTry = 'gemini-flash';

    const promptText = `From the following text, find the purchase order number. It may be labeled as "EBELN", "Pedido", "Orden de Compra", "Purchase Order", "PO No", or "PO #".

Return ONLY the alphanumeric value of the purchase order number itself, without any prefixes, labels, or explanations. The number might contain letters, numbers, and hyphens.

If no purchase order number is found, return an empty string.

The output must be a simple string.

Text:
"${input}"`;

    const response = await ai.generate({
        model: modelToTry,
        prompt: promptText,
    });

    return response.text.trim();
  }
);
