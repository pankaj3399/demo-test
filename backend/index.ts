const express = require("express");
const cors = require("cors");
import { Request, Response } from "express";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { OpenAI } from "openai";
import * as XLSX from "xlsx";
import { Sheet2JSONOpts } from "xlsx";
import * as mammoth from "mammoth";
import * as dotenv from "dotenv";
import * as path from "path";
import type { ChatCompletionMessageParam } from "openai/resources/chat/completions";

interface CompanyAddress {
  street: string;
  secondaryAddress: string;
  city: string;
  state: string;
  zip: string;
}

interface ExcelData {
  divisionName: string;
  countyName: string;
  companyName: string;
  companyType: string;
  dbaName: string;
  websiteAddress: string;
  companyAddress: CompanyAddress;
  bobVisitDateExcel: string;
  earliestExpertScanDateExcel: string;
  retestExpertScanDateExcel: string;
}

interface OpenAIResponse {
  chatGptCompanyDescription: string;
  chatGptParagraph22: string;
  chatGptParagraph19: string;
  nexusFacts40: string;
  section33: string;
}

interface CombinedData extends ExcelData, OpenAIResponse {
  emailSentDate: string;
  date: string;
  region: string;
  division: string;
  county: string;
  section33Docs: string;
  section35Docs: string;
}

interface GeneratePDFRequest extends Request {
  body: {
    url: string;
    emailSentDate: string;
    date: string;
    region: string;
    division: string;
    county: string;
    isRetest: boolean;
  };
}

interface MammothOptions {
  path: string;
  styleMap: string[];
}

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

const extractDataFromExcel = (
  url: string,
  requestDivision: string,
  requestCounty: string,
): ExcelData => {
  const excelFilePath = path.resolve(__dirname, "column_titles.xlsx");
  const workbook = XLSX.readFile(excelFilePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(worksheet) as Record<string, any>[];

  const matchedRow = data.find(
    (row) => row["URL"] && row["URL"].trim() === url.trim(),
  );

  if (!matchedRow) {
    throw new Error(`No data found for URL: ${url}`);
  }

  const excelSerialToDate = (serial: number): string => {
    if (!serial) return "";
    const date = new Date((serial - 25569) * 86400 * 1000);
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };

  return {
    divisionName: requestDivision || "",
    countyName: requestCounty || "",
    companyName: matchedRow["Company_Name"] || "",
    companyType: matchedRow["Type_of_company"] || "",
    dbaName: matchedRow["DBA_Name"] || "",
    websiteAddress: matchedRow["URL"] || "",
    companyAddress: {
      street: matchedRow["Street_Address"] || "",
      secondaryAddress: matchedRow["Secondary_Address"] || "",
      city: matchedRow["City"] || "",
      state: matchedRow["State"] || "",
      zip: matchedRow["ZIP"] || "",
    },
    bobVisitDateExcel: excelSerialToDate(matchedRow["Bob_Visit_Date"] || ""),
    earliestExpertScanDateExcel: excelSerialToDate(
      matchedRow["earliest_Expert_scan_date"] || "",
    ),
    retestExpertScanDateExcel: excelSerialToDate(
      matchedRow["retest_expert_scan_date"] || "",
    ),
  };
};

const processOpenAIPrompts = async (
  url: string,
  companyName: string,
): Promise<OpenAIResponse> => {
  const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
  });

  const prompts = [
    `Define this business and also tell me more about the entity business structure and name ${url}`,
    `Describe this business in detail but in only 2–3 brief but informative sentences. ${url}`,
    `Looking at that website, draft a paragraph explaining what the nexus is between the website and the principal address listed on the website using Defendant to describe "${companyName}" and Defendant's Website to refer to the website.`,
    `Take this paragraph and make it applicable to searching and utilizing the website using Defendant to refer to "${companyName}" and Defendant's Website to refer to the website: "The opportunity to shop and pre-shop Defendant's merchandise available for purchase in the Premises and sign up for an electronic emailer to receive offers, benefits, exclusive invitations, and discounts for use in the Premises from his home are important accommodations for MYERS because traveling outside of his home as a visually disabled individual is often difficult, hazardous, frightening, frustrating, and confusing experience. Defendant has not provided its business information in any other digital format that is accessible for use by blind and visually impaired individuals using the screen reader software."`,
    `Using the information known, finish this paragraph but also include everytime to the reponse: "There is a physical nexus between Defendant's website and Defendant's Premises in that the website provides the contact information, operating hours, and access to products found at Defendant's Premises and address to Defendant's Premises. The website acts as the digital extension of the Principal Place of Business providing ...."`,
  ];

  const messages: ChatCompletionMessageParam[] = [
    {
      role: "system",
      content: `You are analyzing the website ${url} for ${companyName}. Maintain context between responses and refer back to previous information when relevant.`,
    },
  ];

  const results: string[] = [];

  try {
    for (const prompt of prompts) {
      messages.push({
        role: "user",
        content: prompt,
      });

      const completion = await openai.chat.completions.create({
        model: "gpt-4",
        messages: messages,
        temperature: 0.7,
        max_tokens: 500,
      });

      const response = completion.choices[0]?.message?.content || "";
      results.push(response);

      messages.push({
        role: "assistant",
        content: response,
      });
    }

    return {
      chatGptCompanyDescription: results[0],
      chatGptParagraph22: results[1],
      chatGptParagraph19: results[2],
      nexusFacts40: results[3],
      section33: results[4],
    };
  } catch (error) {
    console.error("OpenAI API Error:", error);
    throw new Error(
      error instanceof Error
        ? `Failed to process OpenAI prompts: ${error.message}`
        : "Failed to process OpenAI prompts",
    );
  }
};

const readWordFile = async (filePath: string): Promise<string> => {
  try {
    const options: MammothOptions = {
      path: filePath,
      styleMap: [
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p => p:fresh",
      ],
    };

    const result = await mammoth.extractRawText(options);
    const lines = result.value.split("\n");

    const processedLines = lines.map((line) => {
      const leadingSpaces = line.match(/^[\s\t]*/) || [""];
      const content = line.trim();
      if (content) {
        return `${leadingSpaces[0]}${content}`;
      }
      return "";
    });

    return processedLines.join("\n");
  } catch (error) {
    console.error("Error reading Word file:", error);
    throw error;
  }
};

const formatDate = (date: string | Date | undefined): string => {
  if (!date) return "";

  try {
    let d: Date | undefined;

    if (typeof date === "string") {
      if (!isNaN(Number(date)) && !date.includes("/") && !date.includes("-")) {
        d = new Date((Number(date) - 25569) * 86400 * 1000);
      } else {
        const formats = [
          (str: string) => new Date(str),
          (str: string) => {
            const [day, month, year] = str.split(/[/-]/);
            return new Date(Number(year), Number(month) - 1, Number(day));
          },
          (str: string) => {
            const [month, day, year] = str.split(/[/-]/);
            return new Date(Number(year), Number(month) - 1, Number(day));
          },
          (str: string) => {
            const [year, month, day] = str.split("-");
            return new Date(Number(year), Number(month) - 1, Number(day));
          },
        ];

        for (const format of formats) {
          try {
            const tempDate = format(date);
            if (!isNaN(tempDate.getTime())) {
              d = tempDate;
              break;
            }
          } catch {
            continue;
          }
        }
      }
    } else if (date instanceof Date) {
      d = date;
    }

    if (!d || isNaN(d.getTime())) {
      console.log(`Invalid date value: ${date}`);
      return "";
    }

    const months = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];

    return `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
  } catch (error) {
    console.error("Date formatting error:", error, "for date:", date);
    return "";
  }
};

const generatePDFDocument = async (
  data: CombinedData,
  isRetest = false,
): Promise<Uint8Array> => {
  const pdfDoc = await PDFDocument.create();
  const regularFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const pageSize = { width: 612, height: 792 };
  const leftMargin = 72;
  const rightMargin = 72;
  const topMargin = 72;
  const bottomMargin = 72;

  const fontSize = 11;
  const headerSize = 13;
  const lineHeight = 14;
  const paragraphSpacing = 12;
  const sectionSpacing = 20;
  const maxWidth = pageSize.width - (leftMargin + rightMargin);

  let currentPage = pdfDoc.addPage([pageSize.width, pageSize.height]);
  let yPosition = pageSize.height - topMargin;

  const addNewPage = () => {
    currentPage = pdfDoc.addPage([pageSize.width, pageSize.height]);
    yPosition = pageSize.height - topMargin;
    return currentPage;
  };

  interface DrawTextOptions {
    font?: typeof regularFont;
    size?: number;
    color?: ReturnType<typeof rgb>;
    preserveFormatting?: boolean;
  }

  const drawText = (text: string, options: DrawTextOptions = {}) => {
    if (!text) return;

    const {
      font = regularFont,
      size = fontSize,
      color = rgb(0, 0, 0),
      preserveFormatting = false,
    } = options;

    const wrapText = (textToWrap: string): string[] => {
      const words = textToWrap.split(" ");
      let currentLine = words[0];
      const lines: string[] = [];

      for (let i = 1; i < words.length; i++) {
        const testLine = `${currentLine} ${words[i]}`;
        const width = font.widthOfTextAtSize(testLine, size);

        if (width <= maxWidth) {
          currentLine = testLine;
        } else {
          lines.push(currentLine);
          currentLine = words[i];
        }
      }
      lines.push(currentLine);
      return lines;
    };

    if (preserveFormatting) {
      const lines = text.split("\n");
      for (const line of lines) {
        if (yPosition - lineHeight < bottomMargin) {
          currentPage = addNewPage();
        }

        if (line.trim()) {
          const indentation = (line.match(/^\s*/)?.[0] || "").length * 5;
          const wrappedLines = wrapText(line.trim());

          wrappedLines.forEach((wrappedLine) => {
            currentPage.drawText(wrappedLine, {
              x: leftMargin + indentation,
              y: yPosition,
              size,
              font,
              color,
            });
            yPosition -= lineHeight;
          });
        } else {
          yPosition -= lineHeight / 2;
        }
      }
    } else {
      const wrappedLines = wrapText(text);
      wrappedLines.forEach((line) => {
        if (yPosition - lineHeight < bottomMargin) {
          currentPage = addNewPage();
        }

        currentPage.drawText(line, {
          x: leftMargin,
          y: yPosition,
          size,
          font,
          color,
        });
        yPosition -= lineHeight;
      });
    }
  };

  drawText("Business Analysis Report", {
    size: headerSize + 2,
    font: boldFont,
    color: rgb(0, 0, 0),
  });
  yPosition -= sectionSpacing;

  const drawSection = (label: string, value: string) => {
    if (!value) return;
    drawText(`${label}: ${value}`, { font: regularFont });
    yPosition -= lineHeight / 2;
  };

  drawSection("(Division Name)", data.divisionName);
  drawSection("(County Name)", data.countyName);
  drawSection("(Company Name)", data.companyName);
  drawSection("(Type of Company)", data.companyType);
  drawSection("(DBA Name)", data.dbaName);
  drawSection("(URL)", data.websiteAddress);

  const address = [
    data.companyAddress.street,
    data.companyAddress.secondaryAddress,
    `${data.companyAddress.city}, ${data.companyAddress.state} ${data.companyAddress.zip}`,
  ]
    .filter(Boolean)
    .join(", ");
  drawSection("(Company Address)", address);

  yPosition -= lineHeight / 2;
  drawSection("(Datepicker)", data.date);
  drawSection("(Email Sent Date)", data.emailSentDate);
  drawSection("(Bob Visit Date)", formatDate(data.bobVisitDateExcel));
  drawSection(
    "(Earliest Expert Scan Date)",
    formatDate(data.earliestExpertScanDateExcel),
  );

  if (isRetest) {
    drawSection(
      "(Retest Expert Scan Date)",
      formatDate(data.retestExpertScanDateExcel),
    );
  }

  const drawAnalysisSection = (title: string, content: string) => {
    if (!content) return;
    yPosition -= sectionSpacing;
    drawText(title, {
      font: boldFont,
      size: headerSize,
    });
    yPosition -= lineHeight;

    const paragraphs = content
      .split(/\n{2,}|\r\n{2,}/)
      .map((p) => p.trim())
      .filter((p) => p);
    paragraphs.forEach((paragraph, index) => {
      drawText(paragraph);
      if (index < paragraphs.length - 1) {
        yPosition -= paragraphSpacing;
      }
    });
  };

  // drawAnalysisSection("Company Description", data.chatGptCompanyDescription);
  drawAnalysisSection("( ChatGPTCompanyDescription )", data.chatGptParagraph22);
  drawAnalysisSection("( ChatGPT_Paragraph_19 ) ", data.chatGptParagraph19);
  drawAnalysisSection("( ChatGPT_Paragraph_22 )", data.nexusFacts40);
  drawAnalysisSection("( NexusFacts40 )", data.section33);

  addNewPage();
  const sectionTitle = "Section 33";
  const sectionContent = data.section35Docs;

  if (sectionContent) {
    drawText(sectionTitle, {
      font: boldFont,
      size: headerSize,
    });
    yPosition -= lineHeight * 2;

    drawText(sectionContent, {
      preserveFormatting: true,
    });
  }
  if (isRetest) {
    const sectionTitle = "Section 35";
    const sectionContent = data.section35Docs;
    if (sectionContent) {
      drawText(sectionTitle, {
        font: boldFont,
        size: headerSize,
      });
      yPosition -= lineHeight * 2;

      drawText(sectionContent, {
        preserveFormatting: true,
      });
    }
  }

  return await pdfDoc.save();
};

app.get("/api/urls", (_req: Request, res: Response) => {
  try {
    const filePath = path.join(__dirname, "column_titles.xlsx");
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const opts: Sheet2JSONOpts = { raw: false };
    const data = XLSX.utils.sheet_to_json(worksheet, opts) as Record<
      string,
      any
    >[];

    const urls = data
      .map((row) => row["URL"])
      .filter((URL): URL is string => !!URL);

    res.json(urls);
  } catch (error) {
    console.error("Failed to load URLs:", error);
    res.status(500).json({ error: "Failed to load URLs" });
  }
});

app.post(
  "/api/pdf/generate",
  async (req: GeneratePDFRequest, res: Response) => {
    try {
      const { url, emailSentDate, date, region, division, county, isRetest } =
        req.body;

      const excelData = extractDataFromExcel(url, division, county);
      const openAIData = await processOpenAIPrompts(
        url,
        excelData.companyName || "",
      );

      const wordFilePath = isRetest
        ? path.resolve(__dirname, "Section35.docx")
        : path.resolve(__dirname, "Section33.docx");

      const wordContent = await readWordFile(wordFilePath);

      const combinedData: CombinedData = {
        ...excelData,
        emailSentDate: formatDate(emailSentDate),
        date: formatDate(date),
        region,
        division,
        county,
        ...openAIData,
        section33Docs: wordContent,
        section35Docs: wordContent,
      };

      const pdfBuffer = await generatePDFDocument(combinedData, isRetest);

      res.contentType("application/pdf");
      res.send(Buffer.from(pdfBuffer));
    } catch (error) {
      console.error("PDF Generation Error:", error);
      res.status(500).json({
        error: "Failed to generate PDF",
        message: error instanceof Error ? error.message : "Unknown error",
      });
    }
  },
);

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
