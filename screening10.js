const axios = require('axios');
const pdfParse = require('pdf-parse');
const XLSX = require('xlsx');
const GoogleGenerativeAI = require("@google/generative-ai").GoogleGenerativeAI;
const fs = require('fs');
const path = require('path');

// Function to extract data from the Excel sheet (assuming each row can have the resume URL in a specified column)
function extractUrlsFromExcel(filePath, sheetName) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Read the sheet as a 2D array
        return data; // Return all rows of the sheet
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return [];
    }
}

// Function to convert Google Drive link to exportable link
function convertToExportUrl(googleDriveUrl) {
    const regex = /https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)\/view\?usp=[a-zA-Z0-9_&=-]+/;
    const match = googleDriveUrl.match(regex);
    if (match && match[1]) {
        return `https://drive.google.com/uc?id=${match[1]}`; // Export URL format
    }
    return null; // Return null if the URL is not a valid Google Drive link
}

// Function to fetch the PDF content directly from a URL (public or hosted PDF)
async function fetchPdfFromUrl(url) {
    try {
        const response = await axios.get(url, { responseType: 'arraybuffer' });
        if (response.status === 200) {
            return response.data; // PDF content as a buffer
        } else {
            throw new Error(`Failed to fetch PDF, status code: ${response.status}`);
        }
    } catch (error) {
        console.error(`Error fetching PDF from URL ${url}:`, error.message);
        return null; // Return null if PDF cannot be fetched
    }
}

// Function to extract text from the PDF buffer
async function extractTextFromPDF(pdfBuffer) {
    try {
        const data = await pdfParse(pdfBuffer);
        if (data.text.trim() === '') {
            throw new Error('No text extracted from PDF. The document may be empty or corrupt.');
        }
        return data.text; // Return the extracted text from the PDF
    } catch (error) {
        console.error('Error extracting text from PDF:', error.message);
        return null; // Return null if text cannot be extracted
    }
}

// Function to send extracted resume text and job description to Gemini AI for analysis
async function sendToGeminiAI(resumeText, jobDescriptionText) {
    try {
        const genAI = new GoogleGenerativeAI('AIzaSyDP618tzUGC6m_ceXe48pEQ2PtToqkbKnQ'); // Use environment variable for API key
        const model = genAI.getGenerativeModel({ model: "gemini-pro" });

        const prompt = `I provide you resume and job description, you need to review the resume as ats or hr, match the keywords on the basis of job description and provide just a rating from 10 and return just the number, match keywords and focus on job description and requirements (Skills, Experience in the required role, Education align with role and projects), you are hiring for this role. Resume: ${resumeText} and Job Description: ${jobDescriptionText}`;

        const result = await model.generateContent(prompt);
        const response = await result.response;
        const newText = await response.text(); // Await the text response
        return newText; // Return the generated result
    } catch (error) {
        console.error('Error sending data to Gemini AI:', error.message);
        return null; // Return null if AI analysis fails
    }
}

// Function to update the Excel sheet with results (AI response can go into a specified column)
function updateExcelWithResults(filePath, sheetName, row, resultColumn, analysisResult) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];

        // If no analysis result, insert '0'
        const resultValue = analysisResult ? analysisResult : '0'; // Use '0' if no result or error occurs

        // Convert numerical column index (1-based) to Excel column letter
        const columnLetter = String.fromCharCode(64 + resultColumn); // Converts 13 to 'M', etc.
        const cell = `${columnLetter}${row}`; // Concatenate column letter and row number (e.g., "M2")

        // Write the result to the corresponding cell in the result column
        sheet[cell] = { t: 's', v: resultValue };

        // Save the modified workbook without altering the original formatting
        XLSX.writeFile(workbook, filePath);
    } catch (error) {
        console.error('Error updating Excel file:', error.message);
    }
}

// Helper function to add a small delay (in milliseconds)
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Main function to process resumes and update results in Excel sheet
async function processResumesInExcel(filePath, sheetName, resumeColumn, resultColumn, jobDescriptionUrl) {
    const rows = extractUrlsFromExcel(filePath, sheetName); // Extract all rows from the sheet
    if (rows.length === 0) {
        console.log('No data found in the Excel sheet. Exiting...');
        return;
    }

    // Fetch job description from the provided URL
    let jobDescriptionText = '';
    try {
        const jobDescriptionBuffer = await fetchPdfFromUrl(jobDescriptionUrl);
        if (jobDescriptionBuffer) {
            jobDescriptionText = await extractTextFromPDF(jobDescriptionBuffer);
        }
    } catch (error) {
        console.error('Error fetching or processing job description PDF:', error.message);
        return; // Stop execution if job description PDF fails to load
    }

    if (!jobDescriptionText) {
        console.log('No job description text found. Exiting...');
        return; // Stop execution if no job description is available
    }

    // Loop through each row to process the resume URLs
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];

        // Get the resume URL from the specified column (convert the column letter to index)
        const resumeUrl = row[resumeColumn];
        if (!resumeUrl) {
            console.log(`Skipping row ${rowIndex + 2}: No URL found in column ${resumeColumn}`);
            updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, '0');
            continue; // Skip row if no URL is found
        }

        console.log(`Row ${rowIndex + 2}: Checking URL - ${resumeUrl}`);

        // Convert Google Drive URL to exportable format if it's a valid Google Drive URL
        const exportableResumeUrl = convertToExportUrl(resumeUrl);
        if (!exportableResumeUrl) {
            console.log(`Skipping row ${rowIndex + 2}: Invalid or non-Google Drive URL`);
            updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, '0');
            continue;
        }

        console.log(`Processing resume at URL: ${exportableResumeUrl}`);

        try {
            // Add a small delay to account for loading time
            await delay(2000); // Delay 2 seconds (adjust as needed)

            // Fetch PDF content from the URL (resume)
            const pdfBuffer = await fetchPdfFromUrl(exportableResumeUrl);

            if (!pdfBuffer) {
                console.log(`Skipping row ${rowIndex + 2}: PDF fetch failed for URL: ${exportableResumeUrl}`);
                updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, '0');
                continue; // Skip the row if PDF is not accessible
            }

            let resumeText = await extractTextFromPDF(pdfBuffer);

            if (resumeText && jobDescriptionText) {
                const analysisResult = await sendToGeminiAI(resumeText, jobDescriptionText);
                if (analysisResult) {
                    updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, analysisResult);
                } else {
                    console.log(`No analysis result received for row ${rowIndex + 2}`);
                    updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, '0');
                }
            } else {
                console.log(`No text extracted from resume for URL: ${exportableResumeUrl}`);
                updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, '0');
            }
        } catch (error) {
            console.error(`Error processing PDF for row ${rowIndex + 2}:`, error.message);
            updateExcelWithResults(filePath, sheetName, rowIndex + 2, resultColumn, '0');
        }
    }
}

// Provide the file path, sheet name, columns for resume URL and result, and job description URL
const filePath = 'Aurjobs_Business_analyst_Hiring.xlsx';  // **REPLACE THIS** with the path to your local Excel file
const sheetName = 'Form_Responses_1';  // **REPLACE THIS** with the name of the sheet containing your data (e.g., 'Sheet1')
const resumeColumn = 7;  // **REPLACE THIS** with the column index for resume URLs (9 for column 'J')
const resultColumn = 12;  // **REPLACE THIS** with the column index for results (11 for column 'K')
const jobDescriptionUrl = 'https://drive.google.com/uc?export=download&id=1WPhZX6qxL3Po1o60DrFvivvADcOWDjOe';  // **REPLACE THIS** with your job description PDF URL


// Start processing the Excel URLs
processResumesInExcel(filePath, sheetName, resumeColumn, resultColumn, jobDescriptionUrl).catch((error) => {
    console.error("Error processing resumes:", error.message);
});