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

  const excelSerialToDate = (serial) => {
    if (!serial) return '';
    const date = new Date((serial - 25569) * 86400 * 1000);
    const day = date.getDate();
    const month = date.getMonth() + 1; // ]
    const year = date.getFullYear();
    // Return in DD/MM/YYYY format
    return `${day}/${month}/${year}`;
  };

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
    bobVisitDateExcel: excelSerialToDate(matchedRow['Bob_Visit_Date'] || ''),
    earliestExpertScanDateExcel: excelSerialToDate(matchedRow['earliest_Expert_scan_date'] || '')
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
        model: "gpt-4",
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

const readWordFile = async (filePath) => {
  try {
    // Use mammoth with options to preserve formatting
    const result = await mammoth.extractRawText({
      path: filePath,
      preserveNumbering: true,
      styleMap: [
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p => p:fresh"
      ]
    });
    
    // Split by newlines and preserve formatting
    const lines = result.value.split('\n');
    
    // Process each line to maintain proper spacing and indentation
    const processedLines = lines.map(line => {
      // Preserve leading spaces/tabs for indentation
      const leadingSpaces = line.match(/^[\s\t]*/)[0];
      const content = line.trim();
      if (content) {
        return `${leadingSpaces}${content}`;
      }
      return '';
    });
    
    // Join lines with proper spacing
    return processedLines.join('\n');
  } catch (error) {
    console.error('Error reading Word file:', error);
    throw error;
  }
};

const formatDate = (date) => {
  if (!date) return '';
  
  try {
    let d;
    if (typeof date === 'string') {
      // First try parsing as Excel serial number
      if (!isNaN(date) && !date.includes('/') && !date.includes('-')) {
        d = new Date((Number(date) - 25569) * 86400 * 1000);
      } else {
        // Try multiple date formats
        const formats = [
          // Standard date string
          str => new Date(str),
          // DD/MM/YYYY
          str => {
            const [day, month, year] = str.split(/[/-]/);
            return new Date(year, month - 1, day);
          },
          // MM/DD/YYYY
          str => {
            const [month, day, year] = str.split(/[/-]/);
            return new Date(year, month - 1, day);
          },
          // YYYY-MM-DD
          str => {
            const [year, month, day] = str.split('-');
            return new Date(year, month - 1, day);
          }
        ];

        // Try each format until we get a valid date
        for (const format of formats) {
          try {
            const tempDate = format(date);
            if (!isNaN(tempDate.getTime())) {
              d = tempDate;
              break;
            }
          } catch (e) {
            continue;
          }
        }
      }
    } else if (date instanceof Date) {
      d = date;
    }

    if (!d || isNaN(d.getTime())) {
      console.log(`Invalid date value: ${date}`);
      return '';
    }
    
    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    return `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
  } catch (error) {
    console.error('Date formatting error:', error, 'for date:', date);
    return '';
  }
};

const generatePDFDocument = async (data, isRetest = false) => {
  const pdfDoc = await PDFDocument.create();
  const regularFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const pageSize = { width: 612, height: 792 };
  const leftMargin = 72;    // Adjusted left margin (1 inch)
  const rightMargin = 72;   // Adjusted right margin (1 inch)
  const topMargin = 72;     // Adjusted top margin (1 inch)
  const bottomMargin = 72;  // Adjusted bottom margin (1 inch)
  
  const fontSize = 11;
  const headerSize = 13;
  const lineHeight = 14;    // Reduced line height
  const paragraphSpacing = 12; // Reduced paragraph spacing
  const sectionSpacing = 20;  // Reduced section spacing
  const maxWidth = pageSize.width - (leftMargin + rightMargin);

  let currentPage = pdfDoc.addPage([pageSize.width, pageSize.height]);
  let yPosition = pageSize.height - topMargin;

  const addNewPage = () => {
    currentPage = pdfDoc.addPage([pageSize.width, pageSize.height]);
    yPosition = pageSize.height - topMargin;
    return currentPage;
  };

  const drawText = (text, options = {}) => {
    if (!text) return;

    const {
      font = regularFont,
      size = fontSize,
      color = rgb(0, 0, 0),
      preserveFormatting = false
    } = options;

    // Handle word wrapping for lines that exceed the right margin
    const wrapText = (textToWrap) => {
      const words = textToWrap.split(' ');
      let currentLine = words[0];
      const lines = [];

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
      // Handle formatted text (like Word document content)
      const lines = text.split('\n');
      for (const line of lines) {
        if (yPosition - lineHeight < bottomMargin) {
          currentPage = addNewPage();
        }

        if (line.trim()) {
          const indentation = line.match(/^\s*/)[0].length * 5;
          const wrappedLines = wrapText(line.trim());
          
          wrappedLines.forEach((wrappedLine) => {
            currentPage.drawText(wrappedLine, {
              x: leftMargin + indentation,
              y: yPosition,
              size,
              font,
              color
            });
            yPosition -= lineHeight;
          });
        } else {
          yPosition -= lineHeight / 2; // Reduced spacing for empty lines
        }
      }
    } else {
      // Handle regular text
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
          color
        });
        yPosition -= lineHeight;
      });
    }
  };

  // Document Title
  drawText('Business Analysis Report', {
    size: headerSize + 2,
    font: boldFont,
    color: rgb(0, 0, 0)
  });
  yPosition -= sectionSpacing;

  // Basic Information
  const drawSection = (label, value) => {
    if (!value) return;
    drawText(`${label}: ${value}`, { font: regularFont });
    yPosition -= lineHeight / 2;
  };

  drawSection('Division', data.divisionName);
  drawSection('County', data.countyName);
  drawSection('Company Name', data.companyName);
  drawSection('Type of Company', data.companyType);
  drawSection('DBA Name', data.dbaName);
  drawSection('Website', data.websiteAddress);

  // Address
  const address = [
    data.companyAddress.street,
    data.companyAddress.secondaryAddress,
    `${data.companyAddress.city}, ${data.companyAddress.state} ${data.companyAddress.zip}`
  ].filter(Boolean).join(', ');
  drawSection('Address', address);

  
  yPosition -= lineHeight / 2;
  drawSection('Email Sent', data.emailSentDate);
  drawSection('Bob Visit Date', data.bobVisitDateExcel); // Use raw Excel date
  drawSection('Expert Scan', data.earliestExpertScanDateExcel); // Use raw Excel date

  // Analysis Sections
  const drawAnalysisSection = (title, content) => {
    if (!content) return;
    yPosition -= sectionSpacing;
    drawText(title, {
      font: boldFont,
      size: headerSize
    });
    yPosition -= lineHeight;

    // Split content into paragraphs and handle each with proper spacing
    const paragraphs = content.split(/\n{2,}|\r\n{2,}/).map(p => p.trim()).filter(p => p);
    paragraphs.forEach((paragraph, index) => {
      drawText(paragraph);
      if (index < paragraphs.length - 1) {
        yPosition -= paragraphSpacing;
      }
    });
  };

  drawAnalysisSection('Company Description', data.chatGptCompanyDescription);
  drawAnalysisSection('Business Summary', data.chatGptParagraph22);
  drawAnalysisSection('Website Analysis', data.chatGptParagraph19);
  drawAnalysisSection('Accessibility Assessment', data.nexusFacts40);
  drawAnalysisSection('Physical Location Analysis', data.section33);

  // Word Document Content
  addNewPage();
  const sectionTitle = isRetest ? 'Section 35' : 'Section 33';
  const sectionContent = isRetest ? data.section35Docs : data.section33Docs;

  if (sectionContent) {
    drawText(sectionTitle, {
      font: boldFont,
      size: headerSize
    });
    yPosition -= lineHeight * 2;
    
    drawText(sectionContent, {
      preserveFormatting: true
    });
  }

  return await pdfDoc.save();
};

app.get('/api/urls', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'column_titles.xlsx');
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { 
      raw: false,
      cellDates: true 
    });
    const urls = data
      .map(row => row['URL'])
      .filter(URL => !!URL);

    res.json(urls);
  } catch (error) {
    console.error('Failed to load URLs:', error);
    res.status(500).json({ error: 'Failed to load URLs' });
  }
});



app.post('/api/pdf/generate', async (req, res) => {
  try {
    const { url, emailSentDate, date, region, division, county, isRetest } = req.body;

    const excelData = extractDataFromExcel(url, division, county);
    const openAIData = await processOpenAIPrompts(url, excelData.companyName || '');

    const wordFilePath = isRetest 
      ? path.resolve(__dirname, 'Section35.docx')
      : path.resolve(__dirname, 'Section33.docx');
    
    const wordContent = await readWordFile(wordFilePath);

    const combinedData = {
      ...excelData,  // This will keep bobVisitDateExcel and earliestExpertScanDateExcel as is
      emailSentDate: formatDate(emailSentDate), // Only format the manually entered dates
      date: formatDate(date),
      region,
      division,
      county,
      ...openAIData,
      section33Docs: wordContent,
      section35Docs: wordContent
    };

    const pdfBuffer = await generatePDFDocument(combinedData, isRetest);

    res.contentType('application/pdf');
    res.send(Buffer.from(pdfBuffer));
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
