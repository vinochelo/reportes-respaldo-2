'use server';

/**
 * @fileOverview Extracts the EBELN value from each page of a PDF.
 *
 * - extractEbelnFromPdfPages - A function that handles the extraction process.
 * - ExtractEbelnFromPdfPagesInput - The input type for the extractEbelnFromPdfPages function.
 * - ExtractEbelnFromPdfPagesOutput - The return type for the extractEbelnFromPdfPages function.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const ExtractEbelnFromPdfPagesInputSchema = z.object({
  pdfPages: z.array(
    z.object({
      pageNumber: z.number().describe('The page number of the PDF.'),
      pageText: z.string().describe('The text content of the PDF page.'),
    })
  ).describe('An array of PDF pages with their text content and page numbers.'),
});
export type ExtractEbelnFromPdfPagesInput = z.infer<typeof ExtractEbelnFromPdfPagesInputSchema>;

const ExtractEbelnFromPdfPagesOutputSchema = z.array(
  z.object({
    pageNumber: z.number().describe('The page number of the PDF.'),
    ebeln: z.string().describe('The EBELN value extracted from the page.'),
  })
).describe('An array of page numbers and their corresponding EBELN values.');
export type ExtractEbelnFromPdfPagesOutput = z.infer<typeof ExtractEbelnFromPdfPagesOutputSchema>;

export async function extractEbelnFromPdfPages(
  input: ExtractEbelnFromPdfPagesInput
): Promise<ExtractEbelnFromPdfPagesOutput> {
  return extractEbelnFromPdfPagesFlow(input);
}

const extractEbelnPrompt = ai.definePrompt({
  name: 'extractEbelnPrompt',
  input: {schema: ExtractEbelnFromPdfPagesInputSchema},
  output: {schema: ExtractEbelnFromPdfPagesOutputSchema},
  prompt: `You are an expert data extraction specialist.

  Given the content of each page in a PDF, extract the EBELN value from each page.
  Return an array of JSON objects, where each object contains the pageNumber and the extracted ebeln from that page.
  If no EBELN is found return null.

  Here's an example of the expected output format:
  \`\`\`
  [
    {
      "pageNumber": 1,
      "ebeln": "4500000001"
    },
    {
      "pageNumber": 2,
      "ebeln": "4500000002"
    }
  ]
  \`\`\`

  Here are the PDF pages:
  {{#each pdfPages}}
  Page Number: {{{pageNumber}}}
  Page Text: {{{pageText}}}
  {{/each}}`,
});

const extractEbelnFromPdfPagesFlow = ai.defineFlow(
  {
    name: 'extractEbelnFromPdfPagesFlow',
    inputSchema: ExtractEbelnFromPdfPagesInputSchema,
    outputSchema: ExtractEbelnFromPdfPagesOutputSchema,
  },
  async input => {
    const {output} = await extractEbelnPrompt(input);
    return output!;
  }
);
