const express = require('express');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const OpenAI = require('openai');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const dotenv = require('dotenv');
const cors = require('cors');
const path = require('path');
const fs = require('fs').promises; 

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

const extractDataFromExcel = (url, requestDivision, requestCounty) => {
  const excelFilePath = path.resolve(__dirname, 'column_titles.xlsx');
  const workbook = XLSX.readFile(excelFilePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(worksheet);

  const matchedRow = data.find((row) => row['URL'] && row['URL'].trim() === url.trim());

  if (!matchedRow) {
    throw new Error(`No data found for URL: ${url}`);
  }

  return {
    divisionName: requestDivision || '',
    countyName: requestCounty || '',
    companyName: matchedRow['Company_Name'] || '',
    companyType: matchedRow['Type_of_company'] || '',
    dbaName: matchedRow['DBA_Name'] || '',
    websiteAddress: matchedRow['URL'] || '',
    companyAddress: {
      street: matchedRow['Street_Address'] || '',
      secondaryAddress: matchedRow['Secondary_Address'] || '',
      city: matchedRow['City'] || '',
      state: matchedRow['State'] || '',
      zip: matchedRow['ZIP'] || ''
    },
    bobVisitDateExcel: matchedRow['Bob_Visit_Date'] || '',
    earliestExpertScanDateExcel: matchedRow['earliest_Expert_scan_date'] || '',
  };
};

const processOpenAIPrompts = async (url, companyName) => {
  const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
  });

  const prompts = [
    `Define this business and also tell me more about the entity business structure and name ${url}`,
    `Describe this business in detail but in only 2–3 brief but informative sentences. ${url}`,
    `Looking at that website, draft a paragraph explaining what the nexus is between the website and the principal address listed on the website using Defendant to describe "${companyName}" and Defendant's Website to refer to the website.`,
    `Take this paragraph and make it applicable to searching and utilizing the website using Defendant to refer to "${companyName}" and Defendant's Website to refer to the website: "The opportunity to shop and pre-shop Defendant's merchandise available for purchase in the Premises and sign up for an electronic emailer to receive offers, benefits, exclusive invitations, and discounts for use in the Premises from his home are important accommodations for MYERS because traveling outside of his home as a visually disabled individual is often difficult, hazardous, frightening, frustrating, and confusing experience. Defendant has not provided its business information in any other digital format that is accessible for use by blind and visually impaired individuals using the screen reader software."`,
    `Using the information known, finish this paragraph but also include: "There is a physical nexus between Defendant's website and Defendant's Premises in that the website provides the contact information, operating hours, and access to products found at Defendant's Premises and address to Defendant's Premises. The website acts as the digital extension of the Principal Place of Business providing ...."`
  ];

  const results = await Promise.all(
    prompts.map(async (prompt) => {
      const completion = await openai.chat.completions.create({
        model: "gpt-4o-mini", 
        messages: [{ role: "user", content: prompt }]
      });
      return completion.choices[0]?.message?.content || '';
    })
  );

  return {
    chatGptCompanyDescription: results[0],
    chatGptParagraph22: results[1],
    chatGptParagraph19: results[2],
    nexusFacts40: results[3],
    section33: results[4]
  };
};

app.get('/api/urls', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'column_titles.xlsx');
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    const urls = data
      .map(row => row['URL'])
      .filter(URL => !!URL);

    res.json(urls);
  } catch (error) {
    console.error('Failed to load URLs:', error);
    res.status(500).json({ error: 'Failed to load URLs' });
  }
});

const generatePDFDocument = async (data , isRetest = false) => {
  const pdfDoc = await PDFDocument.create();
  const regularFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.Helvetica);

  const pageSize = { width: 612, height: 792 }; // Standard US Letter size
  const margin = 50;
  const fontSize = 10;
  const headerSize = 12;
  const lineHeight = 14;
  const maxWidth = pageSize.width - (2 * margin);

  let currentPage = pdfDoc.addPage([pageSize.width, pageSize.height]);
  let yPosition = pageSize.height - margin;

  const addNewPage = () => {
    currentPage = pdfDoc.addPage([pageSize.width, pageSize.height]);
    yPosition = pageSize.height - margin;
    return currentPage;
  };

  const cleanMarkdown = (text) => {
    return text
      .replace(/\*\*(.*?)\*\*/g, '$1')
      .replace(/##\s*(.*?)\n/g, '$1\n')
      .replace(/\n+/g, ' ')
      .trim();
  };

  const wrapText = (text, maxWidth, fontSize) => {
    const words = text.split(' ');
    const lines = [];
    let currentLine = words[0];

    for (let i = 1; i < words.length; i++) {
      const width = regularFont.widthOfTextAtSize(currentLine + ' ' + words[i], fontSize);
      if (width < maxWidth) {
        currentLine += ' ' + words[i];
      } else {
        lines.push(currentLine);
        currentLine = words[i];
      }
    }
    lines.push(currentLine);
    return lines;
  };

  const drawWrappedText = (text, options = {}) => {
    const cleanText = text.replace(/[\r\n]+/g, ' ').trim();
    const words = cleanText.split(' ');
    let currentLine = words[0];
    
    for (let i = 1; i < words.length; i++) {
      const testLine = currentLine + ' ' + words[i];
      const width = regularFont.widthOfTextAtSize(testLine, options.size || fontSize);
      
      if (width < maxWidth) {
        currentLine = testLine;
      } else {
        if (yPosition - lineHeight < margin) {
          currentPage = addNewPage();
        }
        
        currentPage.drawText(currentLine, {
          x: options.x || margin,
          y: yPosition,
          size: options.size || fontSize,
          font: options.font || regularFont,
          color: rgb(0, 0, 0)
        });
        
        yPosition -= lineHeight;
        currentLine = words[i];
      }
    }
    
    currentPage.drawText(currentLine, {
      x: options.x || margin,
      y: yPosition,
      size: options.size || fontSize,
      font: options.font || regularFont,
      color: rgb(0, 0, 0)
    });
    
    yPosition -= lineHeight * 1.5;
  };

  const addSection = (label, value) => {
    if (yPosition - (lineHeight * 3) < margin) {
      addNewPage();
    }

    drawWrappedText(label + ':', { 
      size: headerSize, 
      font: boldFont 
    });
    
    if (value) {
      drawWrappedText(value);
    }
  };

  addSection('Division', data.divisionName);
  addSection('County', data.countyName);
  addSection('Company Name', data.companyName);
  addSection('Type of Company', data.companyType);
  addSection('DBA Name', data.dbaName);
  addSection('URL', data.websiteAddress);

  addSection('Location Details', '');
  addSection('Address', [
    data.companyAddress.street,
    data.companyAddress.secondaryAddress,
    `${data.companyAddress.city}, ${data.companyAddress.state} ${data.companyAddress.zip}`
  ].filter(Boolean).join(', '));

  addSection('Important Dates', '');
  addSection('Bob Visit Date', new Date(data.bobVisitDateExcel).toLocaleDateString());
  addSection('Email Sent', new Date(data.emailSentDate).toLocaleDateString());
  addSection('Expert Scan', new Date(data.earliestExpertScanDateExcel).toLocaleDateString());

  // Analysis Sections
  addSection('Detailed Analysis', '');
  addSection('Company Description', data.chatGptCompanyDescription);
  addSection('Business Summary', data.chatGptParagraph22);
  addSection('Website Analysis', data.chatGptParagraph19);
  addSection('Accessibility Assessment', data.nexusFacts40);
  addSection('Physical Location Analysis', data.section33);

  addNewPage();

  if (isRetest) {
    addSection('Section 35', data.section33Docs);

  } else {
    
    addSection('Section 33', data.section35Docs);
  }
  

  const pdfBytes = await pdfDoc.save();
  return Buffer.from(pdfBytes);
};

app.post('/api/pdf/generate', async (req, res) => {
  try {
    const { url, bobVisitDate, emailSentDate, earliestExpertScanDate, date, region, division, county , isRetest } = req.body;

    const excelData = extractDataFromExcel(url, division, county);
    const openAIData = await processOpenAIPrompts(url, excelData.companyName || '');


   async function readWordFile(filePath) {
  try {
    const buffer = await fs.readFile(filePath);
    const result = await mammoth.extractRawText({ buffer });
    
    return result.value
      .replace(/[\r\n]+/g, ' ')  
      .replace(/\s{2,}/g, ' ') 
      .trim();
  } catch (error) {
    console.error('Error reading Word file:', error);
    throw error;
  }
}
    
    

    const wordFilePath = isRetest 
      ? path.resolve(__dirname, 'Section33.docx')
      : path.resolve(__dirname, 'Section35.docx');
    
    const wordContent = await readWordFile(wordFilePath);

    const combinedData = {
      ...excelData,
      bobVisitDate,
      emailSentDate,
      earliestExpertScanDate,
      date,
      region,
      division,
      county,
      chatGptCompanyDescription: openAIData.chatGptCompanyDescription || '',
      chatGptParagraph22: openAIData.chatGptParagraph22 || '',
      chatGptParagraph19: openAIData.chatGptParagraph19 || '',
      nexusFacts40: openAIData.nexusFacts40 || '',
      section33: openAIData.section33 || '',
      section33Docs: wordContent || '',
      section35Docs: wordContent || ''

    };

    const pdfBuffer = await generatePDFDocument(combinedData , isRetest);

    res.contentType('application/pdf');
    res.send(pdfBuffer);
  } catch (error) {
    console.error('PDF Generation Error:', error);
    res.status(500).json({
      error: 'Failed to generate PDF',
      message: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});