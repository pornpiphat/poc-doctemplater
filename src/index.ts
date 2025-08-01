import * as fs from 'fs';
import { TemplateHandler } from 'easy-template-x';
import path from 'path';
import { convertDocxToPDF } from './shared/services/libreoffice.service';

const handler = new TemplateHandler();

const outputFolder = path.resolve(__dirname, "./output");
// const template1 = async () => {
//     console.log("template1");

//     const templatePath = path.resolve(__dirname, './templates/t001.docx');
//     const templateFile = fs.readFileSync(templatePath);
//     const data = {
//         "name": "Gianni Castiglione",
//         "address": "111 1st Avenue",
//         "city": "Redmond",
//         "state": "WA",
//         "zip": "67891",
//         "recipient": "Yamada",
//         "message": "Dear Mr. Sagese,\n\nI am writing to express my interest in the position at Shakti Bank. With my background and experience, I believe I would be a valuable addition to your team.\n\nThank you for considering my application. I look forward to the opportunity to speak with you further.\n\nSincerely,\nGianni Castiglione",
//         "link": {
//             _type: 'link',
//             text: 'Click here',  // Optional - if not specified the `target` property will be used
//             target: 'https://www.youtube.com/watch?v=Ay7Fg3OwGeQ&list=RD8oMgX7t1ANI&index=4'
//         }
//     }
//     const handler = new TemplateHandler();
//     const docxBuf: NonSharedBuffer = await handler.process(templateFile, data);
//     const pdfBuf = await convertDocxToPDF(docxBuf);


//     fs.writeFileSync(path.join(outputFolder, 'too1-output.pdf'), pdfBuf);
// };

const template2 = async () => {
    const templatePath = path.resolve(__dirname, './templates/t002.docx');
    const templateFile = fs.readFileSync(templatePath);
    const data = {
        companyName: "MOCK Company",
        newsletterTitle: 'Monthly newsletter',
        date: '20 September 2025',
        Title: "Welcome to our September newsletter!",
        content: "This month has been a month of change and innovation in our organization.\n\nWe recently launched a new product, 'SmartX Hub', an IoT platform that connects smart home devices to work together.\n\nCongratulations to the app development team for winning the Best of Innovation award at TechCon Asia 2025 last week.\n\nThis month, we've added 5 new employees to our team. We hope you all give us a warm welcome.\n\nDon't forget to tune in to our 'Friday Talk' event every Friday at 4:00 PM in the 5th floor conference room to stay updated on the latest knowledge in our organization.",
        link: {
            _type: 'link',
            text: ' View Video',
            target: 'https://www.youtube.com/watch?v=ekr2nIex040&list=RDekr2nIex040&start_radio=1'
        },
        image: {
            _type: "image",
            source: fs.readFileSync(path.resolve(__dirname, './images/IT-Company.jpg')),
            format: 'image/jpeg',
            altText: "IT Company",
            width: 150,
            height: 150
        },
        Beers: [
            { Brand: "Carlsberg", Price: 1 },
            { Brand: "Leaf Blonde", Price: 2 },
            { Brand: "Weihenstephan", Price: 1.5 }
        ]
    };
    const docxBuf: NonSharedBuffer = await handler.process(templateFile, data);
    const pdfBuf = await convertDocxToPDF(docxBuf);
    // fs.writeFileSync(path.join(outputFolder, 'too2-output.docx'), docxBuf);
    fs.writeFileSync(path.join(outputFolder, 'too2-output.pdf'), pdfBuf);
    console.log("PDF file has been successfully generated from the template!");
}


const main = () => {
    // template1();
    template2();
};

main();

