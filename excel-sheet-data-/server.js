const express = require('express');
const mongoose = require('mongoose');
const expressFileUpload = require('express-fileupload');
const exceljs = require('exceljs');
const path = require('path');
const fs = require('fs');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const ejs = require('ejs');

const app = express();
const port = 3000;

mongoose.connect('mongodb+srv://admin2:admin21234@cluster0.bny17md.mongodb.net/excelData?retryWrites=true&w=majority', { useNewUrlParser: true, useUnifiedTopology: true });

const TreeDonation = mongoose.model('TreeDonation', {
    name: String,
    email: String,
    phoneNumber: String,
    amount: Number,
    treesDonated: Number,
});

app.use(expressFileUpload());
app.use(express.static(path.join(__dirname, 'public')));
app.set('view engine', 'ejs');

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/submit', async (req, res) => {
    try {
        const excelFile = req.files.excelFile;
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.load(excelFile.data);
        const sheet = workbook.getWorksheet(1);

        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: 'harshita0583.be21@chitkara.edu.in',
                pass: 'Harshita2003',
            },
        });

        sheet.eachRow(async (row, rowNumber) => {
            if (rowNumber > 1) {
                const [empty, name, emailObj, phoneNumber, amount] = row.values;
                const email = emailObj.text;
                const treesDonated = Math.floor(amount / 100);

                // Generate PDF certificate
                const certificatesDirectory = path.join(__dirname, 'certificates');
                if (!fs.existsSync(certificatesDirectory)) {
                    fs.mkdirSync(certificatesDirectory);
                }

                const pdfFilePath = path.join(certificatesDirectory, `${name}_certificate.pdf`);
                const pdfStream = fs.createWriteStream(pdfFilePath);
                const pdfDoc = new PDFDocument();
                pdfDoc.pipe(pdfStream);
                pdfDoc.fontSize(20).text(`Certificate of Tree Donation`, { align: 'center' });
                pdfDoc.moveDown(0.5);
                pdfDoc.fontSize(16).text(`Presented to`, { align: 'center' });
                pdfDoc.moveDown(0.5);
                pdfDoc.fontSize(18).text(`${name}`, { align: 'center' });
                pdfDoc.moveDown(1);
                pdfDoc.fontSize(14).text(`In recognition of your generous contribution`, { align: 'center' });
                pdfDoc.moveDown(0.5);
                pdfDoc.fontSize(14).text(`Amount Donated: $${amount}`, { align: 'center' });
                pdfDoc.fontSize(14).text(`Number of Trees Donated: ${treesDonated}`, { align: 'center' });
                pdfDoc.end();

                const templatePath = path.join(__dirname, 'views', 'certificate.ejs');
                const template = fs.readFileSync(templatePath, 'utf-8');
                const renderedHTML = ejs.render(template, { name, amount, treesDonated });

                // Send email with the PDF certificate attached
                const mailOptions = {
                    from: 'harshita0583.be21@chitkara.edu.in',
                    to: email,
                    subject: 'Congratulations on Your Tree Donation!',
                    text: 'Thank you for your generous donation.',
                    // html: renderedHTML,
                    attachments: [
                        {
                            filename: 'certificate.pdf',
                            path: pdfFilePath,
                            encoding: 'base64',
                        },
                    ],
                };

                transporter.sendMail(mailOptions, (error, info) => {
                    if (error) {
                        console.error(error);
                    } else {
                        console.log('Email sent: ' + info.response);
                    }
                });

                await TreeDonation.create({
                    name,
                    email,
                    phoneNumber,
                    amount,
                    treesDonated,
                });
            }
        });

        res.send('Data imported successfully!');
    } catch (error) {
        console.error(error);
        res.status(500).send('Error importing data');
    }
});

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});