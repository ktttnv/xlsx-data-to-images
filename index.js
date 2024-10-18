const reader = require('xlsx')
const fs = require('fs');
const { createCanvas } = require("canvas");
const { PDFDocument } = require('pdf-lib');

const IMAGE_WIDTH_MM = 94;
const IMAGE_HEIGHT_MM = 75;

const ONE_INCH_IN_MM = 25.4;
const ONE_INCH_IN_POINTS = 72;

const A3_WIDTH_MM = 297;
const A3_HEIGHT_MM = 420;

const MARGIN_MM = 5;
const SPACING_MM = 2;

const TABLE_WITH_DATA_PATH = './test.xlsx';
const OUTPUT_FOLDER_PATH = '/_output';
const IMAGES_FOLDER_PATH = OUTPUT_FOLDER_PATH + '/images';
const OUTPUT_PDF_PATH = OUTPUT_FOLDER_PATH + '/output.pdf';

function readXlsxFile(filePath) {
    const file = reader.readFile(filePath);

    const data = [];
    const sheets = file.SheetNames;
    
    for (let i = 0; i < sheets.length; i++) {
        const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
        temp.forEach((res) => {
            data.push({
                surname: res['Фамилия'] ?? "",
                name: res['Имя'] ?? "",
                company: res['Компания или учебное заведение'] ?? "",
                position: res['Должность'] ?? ""

            });
        });
    }

    return data;
}

function createFolder(relativeFolderPath) {
    const folderPath = __dirname + relativeFolderPath;

    try {
        if (!fs.existsSync(folderPath)) {
            fs.mkdirSync(folderPath);
        }
    } catch (err) {
        console.error(err);
    }
}

function removeFolder(relativeFolderPath) {
    const folderPath = __dirname + relativeFolderPath;

    try {
        if (!fs.existsSync(folderPath)) {
            fs.rm(folderPath, { recursive: true, force: true });
        }
    } catch (err) {
        console.error(err);
    }
}

function convertMmToPixels(mm) {
    return (mm / ONE_INCH_IN_MM) * 300;
}

function formatLongString(str) {
    if (str.length < 30) return [str, ""];

    let pos = 0;

    for (let i = Math.floor(str.length / 2); i < str.length; i++) {
        if (str[i] === ' ') {
            pos = i;
            break;
        }
    }

    return [str.slice(0, pos), str.slice(pos + 1, str.length)];
}

function createImage(data, relativeFolderPath, imageName) {
    const folderPath = __dirname + relativeFolderPath;
    
    const width = convertMmToPixels(IMAGE_WIDTH_MM);
    const height = convertMmToPixels(IMAGE_HEIGHT_MM);
    const canvas = createCanvas(width, height);

    const context = canvas.getContext("2d");

    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);

    context.textAlign = "center";
    context.fillStyle = "#000000";

    context.font = "bold 130px Arial";
    context.fillText(data.name.toUpperCase(), width / 2, height / 100 * 30);
   
    context.font = "bold 90px Arial";
    context.fillText(data.surname.toUpperCase(), width / 2, height / 100 * 45);
    
    context.font = "60px Arial";
    companyData = formatLongString(data.company);
    context.fillText(companyData[0], width / 2, height / 100 * 57);
    context.fillText(companyData[1], width / 2, height / 100 * 63);
    
    context.font = "55px Arial";
    positionData = formatLongString(data.position);
    context.fillText(positionData[0], width / 2, height / 100 * 82);
    context.fillText(positionData[1], width / 2, height / 100 * 88);

    // context.strokeStyle = "#808080";
    // context.lineWidth = 2;
    // context.strokeRect(0, 0, width, height);

    const buffer = canvas.toBuffer("image/png");

    try {
        fs.writeFileSync(folderPath + "/" + imageName + ".png", buffer);
    } catch (err) {
        console.error(err);
    }
}

function convertMmToPoints(mm) {
    return mm / ONE_INCH_IN_MM * ONE_INCH_IN_POINTS;
}

async function createPdfWithImages(imagePaths, outputPdfPath) {
    const pdfDoc = await PDFDocument.create();
  
    const pageWidth = convertMmToPoints(A3_WIDTH_MM);
    const pageHeight = convertMmToPoints(A3_HEIGHT_MM);

    const imageWidth = convertMmToPoints(IMAGE_WIDTH_MM);
    const imageHeight = convertMmToPoints(IMAGE_HEIGHT_MM);
  
    const margin = convertMmToPoints(MARGIN_MM);
    const spacing = convertMmToPoints(SPACING_MM);
    
    let currentPage = pdfDoc.addPage([pageWidth, pageHeight]);
    let x = margin;
    let y = pageHeight - margin - imageHeight;
  
    for (const imagePath of imagePaths) {
      const imageBytes = fs.readFileSync(imagePath);
      const image = await pdfDoc.embedPng(imageBytes);
  
      currentPage.drawImage(image, {
        x: x,
        y: y,
        width: imageWidth,
        height: imageHeight,
      });
  
      x += imageWidth + spacing;
  
      if (x + imageWidth > pageWidth - margin) {
        x = margin;
        y -= imageHeight + spacing;
      }
  
      if (y < margin) {
        currentPage = pdfDoc.addPage([pageWidth, pageHeight]);
        x = margin;
        y = pageHeight - margin - imageHeight;
      }
    }
  
    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPdfPath, pdfBytes);
  }

const tableData = readXlsxFile(TABLE_WITH_DATA_PATH);

removeFolder(OUTPUT_FOLDER_PATH);
createFolder(OUTPUT_FOLDER_PATH);
createFolder(IMAGES_FOLDER_PATH);

const imagePaths = [];

tableData.forEach((person, ind) => {
    const imageName = "image" + (ind + 1).toString().padStart(3, '0');
    createImage(person, IMAGES_FOLDER_PATH, imageName);
    const imagePath = __dirname + IMAGES_FOLDER_PATH + "/" + imageName + ".png";
    imagePaths.push(imagePath);
});

createPdfWithImages(imagePaths, __dirname + OUTPUT_PDF_PATH)
  .then(() => console.log('PDF created successfully'))
  .catch(err => console.error('Error while creating PDF:', err));
