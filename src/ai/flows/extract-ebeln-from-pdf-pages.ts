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
  model: googleAI.model('gemini-1.5-flash'),
  input: {schema: ExtractEbelnInputSchema},
  prompt: `You are an expert at finding purchase order numbers in document text. The purchase order number is often labeled as "Pedido", "EBELN", "PO No.", or "Orden de Compra". It can be a purely numeric value (like 4500123456) or alphanumeric. Find the purchase order number in the following text.

Respond with a valid JSON object with a single key "ebeln". The value should be the purchase order number as a string, or null if no number is found. Do not include any other text or markdown formatting in your response.

Example response: {"ebeln": "4500123456"}

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
    const response = await extractEbelnPrompt(input);
    const textOutput = response.text.trim();
    try {
      // The model may wrap the JSON in markdown backticks.
      const jsonText = textOutput.replace(/^```json\n?/, '').replace(/```$/, '');
      const parsed = JSON.parse(jsonText);
      // Validate the parsed object against our schema.
      return ExtractEbelnOutputSchema.parse(parsed);
    } catch (e) {
      console.error("Failed to parse AI JSON response:", textOutput, e);
      // Let the calling function handle the error display.
      throw new Error("AI returned an invalid response format.");
    }
  }
);

export async function extractEbeln(input: ExtractEbelnInput): Promise<ExtractEbelnOutput> {
  return extractEbelnFlow(input);
}
