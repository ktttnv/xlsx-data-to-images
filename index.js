const reader = require('xlsx')
const fs = require('fs');
const { createCanvas } = require("canvas");
const { PDFDocument } = require('pdf-lib');

const IMAGE_WIDTH_MM = 96;
const IMAGE_HEIGHT_MM = 75;

const ONE_INCH_IN_MM = 25.4;
const ONE_INCH_IN_POINTS = 72;

const A3_WIDTH_MM = 297;
const A3_HEIGHT_MM = 420;

const MARGIN_MM = 20;
const SPACING_MM = 20;

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
                surname: res['Фамилия'],
                name: res['Имя'],
                company: res['Компания или учебное заведение'],
                position: res['Должность']

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

function convertMmToPixels(mm) {
    return (mm / ONE_INCH_IN_MM) * 300;
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
    
    context.font = "40px Arial";
    context.fillText(data.company, width / 2, height / 100 * 60);
    
    context.font = "35px Arial";
    context.fillText(data.position, width / 2, height / 100 * 85);

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
  
    const margin = convertMmToPoints(MARGIN_MM);
    const spacing = convertMmToPoints(SPACING_MM);
    
    let currentPage = pdfDoc.addPage([pageWidth, pageHeight]);
    let x = margin;
    let y = pageHeight - margin - IMAGE_HEIGHT_MM;
  
    for (const imagePath of imagePaths) {
      const imageBytes = fs.readFileSync(imagePath);
      const image = await pdfDoc.embedPng(imageBytes);
  
      currentPage.drawImage(image, {
        x: x,
        y: y,
        width: IMAGE_WIDTH_MM,
        height: IMAGE_HEIGHT_MM,
      });
  
      x += IMAGE_WIDTH_MM + spacing;
  
      if (x + IMAGE_WIDTH_MM > pageWidth - margin) {
        x = margin;
        y -= IMAGE_HEIGHT_MM + spacing;
      }
  
      if (y < margin) {
        currentPage = pdfDoc.addPage([pageWidth, pageHeight]);
        x = margin;
        y = pageHeight - margin - IMAGE_HEIGHT_MM;
      }
    }
  
    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPdfPath, pdfBytes);
  }

const tableData = readXlsxFile(TABLE_WITH_DATA_PATH);

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
