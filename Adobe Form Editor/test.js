const { PDFDocument, PDFTextField, PDFRadioGroup } = require('pdf-lib');
const fs = require('fs').promises;

async function updateFields(inputPdfPath, outputPdfPath, fieldNames, fieldValues) {
    try {
        const pdfBytes = await fs.readFile(inputPdfPath);
        const pdfDoc = await PDFDocument.load(pdfBytes);
        const form = pdfDoc.getForm();

        for (let i = 0; i < fieldNames.length; i++) {
            const fieldName = fieldNames[i];
            const fieldValue = fieldValues[i];

            form.getFields().forEach(field => {
                if (field instanceof PDFTextField) {
                    field.setFontSize(10);
                }
            });

            const field = form.getField(fieldName);
            if (!field) {
                console.log(`Field "${fieldName}" not found.`);
                continue;
            }

            if (field instanceof PDFTextField) {
                let transformedValue = fieldValue.replace(/([A-Za-z])(\w*)/g, function(_, firstChar, rest) {
                    return firstChar.toUpperCase() + rest.toLowerCase();
                });

                field.setText(transformedValue);
            } else if (field instanceof PDFRadioGroup) {
                if (fieldValue === '$yes') {
                    field.select(field.getOptions()[0]);
                } else if (fieldValue === '$no') {
                    field.select(field.getOptions()[1]);
                } else {
                    console.log(`Invalid value "${fieldValue}" for radio button group "${fieldName}".`);
                }
            } else {
                console.log(`Unsupported field type for field "${fieldName}".`);
            }
        }

        const modifiedPdfBytes = await pdfDoc.save();
        await fs.writeFile(outputPdfPath, modifiedPdfBytes);

        return 'PDF fields updated successfully.'

    } catch (error) {
        return { error: error.message };
    }
}

async function printFieldNames(inputPdfPath) {
    try {
        const pdfBytes = await fs.readFile(inputPdfPath);
        const pdfDoc = await PDFDocument.load(pdfBytes);
        const form = pdfDoc.getForm();
        const fields = form.getFields();
        const fieldNames = fields.map(field => field.getName());

        return fieldNames;
    } catch (error) {
        return { error: error.message };
    }
}

// Parse command-line arguments
const [, , functionName, ...args] = process.argv;

// Dispatch function based on provided function name
switch (functionName) {
    case 'printFieldNames':
        printFieldNames(...args)
            .then(fieldNames => console.log("Field Names:", fieldNames))
            .catch(error => console.error('Error:', error));
        break;
    case 'updateFields':
        // Extract inputPdfPath and outputPdfPath
        const [inputPdfPath, outputPdfPath, ...fields] = args;
        
        // Split fields array into fieldNames and fieldValues arrays
        const fieldNames = fields.slice(0, fields.length / 2);
        const fieldValues = fields.slice(fields.length / 2);
        
        updateFields(inputPdfPath, outputPdfPath, fieldNames, fieldValues)
            .then(result => console.log(result))
            .catch(error => console.error('Error:', error));
        break;
    default:
        console.error('Unknown function name. Supported functions: printFieldNames, updateFields.');
        break;
}
    