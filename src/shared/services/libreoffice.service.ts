import libre from 'libreoffice-convert';

export const convertDocxToPDF = async (docxBuf: Buffer): Promise<Buffer> => {
    return new Promise((resolve, reject) => {
        libre.convert(docxBuf, '.pdf', undefined, (err, pdfBuf) => {
            if (err || !pdfBuf) return reject(err);
            resolve(pdfBuf);
        });
    });
};
