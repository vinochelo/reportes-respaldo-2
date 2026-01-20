'use server';

/**
 * @fileOverview Extracts the BELNR value from each page of a PDF.
 *
 * - extractBelnrFromPdfPages - A function that handles the extraction process.
 * - ExtractBelnrFromPdfPagesInput - The input type for the extractBelnrFromPdfPages function.
 * - ExtractBelnrFromPdfPagesOutput - The return type for the extractBelnrFromPdfPages function.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const ExtractBelnrFromPdfPagesInputSchema = z.object({
  pdfPages: z.array(
    z.object({
      pageNumber: z.number().describe('The page number of the PDF.'),
      pageText: z.string().describe('The text content of the PDF page.'),
    })
  ).describe('An array of PDF pages with their text content and page numbers.'),
});
export type ExtractBelnrFromPdfPagesInput = z.infer<typeof ExtractBelnrFromPdfPagesInputSchema>;

const ExtractBelnrFromPdfPagesOutputSchema = z.array(
  z.object({
    pageNumber: z.number().describe('The page number of the PDF.'),
    belnr: z.string().describe('The BELNR value extracted from the page.'),
  })
).describe('An array of page numbers and their corresponding BELNR values.');
export type ExtractBelnrFromPdfPagesOutput = z.infer<typeof ExtractBelnrFromPdfPagesOutputSchema>;

export async function extractBelnrFromPdfPages(
  input: ExtractBelnrFromPdfPagesInput
): Promise<ExtractBelnrFromPdfPagesOutput> {
  return extractBelnrFromPdfPagesFlow(input);
}

const extractBelnrPrompt = ai.definePrompt({
  name: 'extractBelnrPrompt',
  input: {schema: ExtractBelnrFromPdfPagesInputSchema},
  output: {schema: ExtractBelnrFromPdfPagesOutputSchema},
  prompt: `You are an expert data extraction specialist.

  Given the content of each page in a PDF, extract the BELNR value from each page.
  Return an array of JSON objects, where each object contains the pageNumber and the extracted belnr from that page.
  If no BELNR is found return null.

  Here's an example of the expected output format:
  [
    {
      "pageNumber": 1,
      "belnr": "5105646951"
    },
    {
      "pageNumber": 2,
      "belnr": "5105646952"
    }
  ]

  Here are the PDF pages:
  {{#each pdfPages}}
  Page Number: {{{pageNumber}}}
  Page Text: {{{pageText}}}
  {{/each}}`,
});

const extractBelnrFromPdfPagesFlow = ai.defineFlow(
  {
    name: 'extractBelnrFromPdfPagesFlow',
    inputSchema: ExtractBelnrFromPdfPagesInputSchema,
    outputSchema: ExtractBelnrFromPdfPagesOutputSchema,
  },
  async input => {
    const {output} = await extractBelnrPrompt(input);
    return output!;
  }
);
