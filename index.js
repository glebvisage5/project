let express = require('express');
const bodyparser = require('body-parser');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const fs2 = require('fs').promises;
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const cors = require('cors'); //
const archiver = require('archiver');

const app = express();
app.use(cors());

// Call middleware, storing a reference to it:
app.use(express.static('uploads'));
app.use(express.static('public'));
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, "uploads");
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});


var upload = multer({ storage: storage });

app.use(bodyparser.urlencoded({ extended: false }));
app.use(bodyparser.json());

app.get('/', (req, res) => {
    res.sendFile(__dirname + "/public/index.html");
});
async function main(inputPath) {
    const ext = '.pdf'
    const outputPath = path.join(__dirname, `/output${ext}`);
    // Read file
    const docxBuf = await fs2.readFile(inputPath);

    // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
    let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);

    // Here in done you have pdf file which you can save or transfer in another stream
    await fs2.writeFile(outputPath, pdfBuf);
    setTimeout(() => {
        fs.unlink("output.pdf", (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            } else {
                console.log(`Файл успешно удален.`);
            }
        });
        fs.unlink("output.docx", (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            } else {
                console.log(`Файл успешно удален.`);
            }
        });
    }, 4 * 60 * 1000); // 4 минуты в миллисекундах
}
async function mainforDop(inputPath, inputPath2) {
    const ext = '.pdf'
    const outputPath = path.join(__dirname, `/output${ext}`);
    // Read file
    const docxBuf = await fs2.readFile(inputPath);

    // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
    let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);

    // Here in done you have pdf file which you can save or transfer in another stream
    await fs2.writeFile(outputPath, pdfBuf);
    const outputPath2 = path.join(__dirname, `/output2${ext}`);
    // Read file
    const docxBuf2 = await fs2.readFile(inputPath2);

    // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
    let pdfBuf2 = await libre.convertAsync(docxBuf2, ext, undefined);

    // Here in done you have pdf file which you can save or transfer in another stream
    await fs2.writeFile(outputPath2, pdfBuf2);
    setTimeout(() => {
        fs.unlink("output.pdf", (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            } else {
                console.log(`Файл успешно удален.`);
            }
        });
        fs.unlink("output.docx", (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            } else {
                console.log(`Файл успешно удален.`);
            }
        });
        fs.unlink("output2.pdf", (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            } else {
                console.log(`Файл успешно удален.`);
            }
        });
        fs.unlink("output2.docx", (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            } else {
                console.log(`Файл успешно удален.`);
            }
        });

    }, 4 * 60 * 1000); // 3 минуты в миллисекундах
}

app.post('/download', upload.single('file'), (req, res) => {
    if (req.body.isPDF) {
        fs.access('output2.pdf', fs.constants.F_OK, (err) => {
            if (err) {
                fs.access('output.pdf', fs.constants.F_OK, (err) => {
                    if (err) {
                        console.error(`Файл не найден.`);
                    } else {
                        res.download('output.pdf', () => {
                            if (err) {
                                console.error('Ошибка при скачивании файла:', err);
                                res.status(500).send('Ошибка при скачивании файла');
                            } else {
                                fs.unlink('output.pdf', (unlinkErr) => {
                                    if (unlinkErr) {
                                        console.error('Ошибка при удалении файла:', unlinkErr);
                                    } else {
                                        console.log(`Файл  успешно удален.`);
                                    }
                                });
                                // Удаляем файл после успешной отправки
                                fs.unlink('output.docx', (unlinkErr) => {
                                    if (unlinkErr) {
                                        console.error('Ошибка при удалении файла:', unlinkErr);
                                    } else {
                                        console.log(`Файл успешно удален.`);
                                    }
                                });
                            }

                        })

                    }
                });
            } else {
                const archive = archiver('zip');
                const zipFileName = 'ШЗ.zip';
                const output = fs.createWriteStream(zipFileName);
                archive.file('output.pdf', { name: 'output.pdf' });
                archive.file('output2.pdf', { name: 'output2.pdf' });
                archive.pipe(output);
                // Завершаем архивацию и отправляем его пользователю
                archive.finalize();


                output.on('close', () => {
                    res.download(zipFileName, (err) => {
                        if (err) {
                            console.error('Ошибка при скачивании архива:', err);
                            res.status(500).send('Ошибка при скачивании архива');
                        } else {
                            // Удаляем zip-архив после успешной отправки
                            fs.unlink(zipFileName, (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении архива:', unlinkErr);
                                } else {
                                    console.log(`Архив ${zipFileName} успешно удален.`);
                                }
                            });
                            fs.unlink('output.pdf', (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении файла:', unlinkErr);
                                } else {
                                    console.log(`Файл  успешно удален.`);
                                }
                            });
                            // Удаляем файл после успешной отправки
                            fs.unlink('output.docx', (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении файла:', unlinkErr);
                                } else {
                                    console.log(`Файл успешно удален.`);
                                }
                            });
                            fs.unlink('output2.pdf', (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении файла:', unlinkErr);
                                } else {
                                    console.log(`Файл  успешно удален.`);
                                }
                            });
                            // Удаляем файл после успешной отправки
                            fs.unlink('output2.docx', (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении файла:', unlinkErr);
                                } else {
                                    console.log(`Файл успешно удален.`);
                                }
                            });

                        }
                    });
                });
            }
        });
    } else {
        fs.access('output2.docx', fs.constants.F_OK, (err) => {
            if (err) {
                fs.access('output.docx', fs.constants.F_OK, (err) => {
                    if (err) {
                        console.error(`Файл не найден.`);
                    } else {
                        res.download('output.docx', () => {
                            if (err) {
                                console.error('Ошибка при скачивании файла:', err);
                                res.status(500).send('Ошибка при скачивании файла');
                            } else {
                                // Удаляем файл после успешной отправки
                                fs.unlink('output.docx', (unlinkErr) => {
                                    if (unlinkErr) {
                                        console.error('Ошибка при удалении файла:', unlinkErr);
                                    } else {
                                        console.log(`Файл  успешно удален.`);
                                    }
                                });
                            }
                        })

                    }
                });
            } else {
                const archive = archiver('zip');
                const zipFileName = 'ШЗ.zip';
                const output = fs.createWriteStream(zipFileName);
                archive.file('output.docx', { name: 'output.docx' });
                archive.file('output2.docx', { name: 'output2.docx' });
                archive.pipe(output);
                // Завершаем архивацию и отправляем его пользователю
                archive.finalize();


                output.on('close', () => {
                    res.download(zipFileName, (err) => {
                        if (err) {
                            console.error('Ошибка при скачивании архива:', err);
                            res.status(500).send('Ошибка при скачивании архива');
                        } else {
                            // Удаляем zip-архив после успешной отправки
                            fs.unlink(zipFileName, (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении архива:', unlinkErr);
                                } else {
                                    console.log(`Архив ${zipFileName} успешно удален.`);
                                }
                            });
                            // Удаляем файл после успешной отправки
                            fs.unlink('output.docx', (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении файла:', unlinkErr);
                                } else {
                                    console.log(`Файл успешно удален.`);
                                }
                            });

                            // Удаляем файл после успешной отправки
                            fs.unlink('output2.docx', (unlinkErr) => {
                                if (unlinkErr) {
                                    console.error('Ошибка при удалении файла:', unlinkErr);
                                } else {
                                    console.log(`Файл успешно удален.`);
                                }
                            });

                        }
                    });
                });
            }
        });
    }

});

app.post('/submit', (req, res) => {
    const info = req.body.info;
    const content = fs.readFileSync(
        path.resolve(__dirname, "example/REFERAT.docx"),
        "binary"
    );
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.render({
        replaceAuthors: info[0] || ' ',
        replaceProg: info[1] || ' ',
        replaceAnot: info[2] || ' ',
        replaceIVM: info[3] || ' ',
        replaceLa: info[4] || ' ',
        replaceOCS: info[5] || ' ',
        replaceOB: info[6] || ' '
    });
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
    main(path.join(__dirname, "output.docx"));

});
app.post('/submit2', (req, res) => {
    const info = req.body.info;
    const content = fs.readFileSync(
        path.resolve(__dirname, "example/adad.docx"),
        "binary"
    );
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.render({
        replaceProduct: info[0] || ' ',
        replaceFIO: info[1] || ' ',
        replaceIndex: info[2] || ' ',
        replaceCountry: info[3] || ' ',
        replaceCity: info[4] || ' ',
        replaceAdress: info[5] || ' ',
        replaceAdressNum: info[6] || ' ',
        replaceCorpus: info[7] || ' ',
        replaceNumbH: info[8] || ' ',
        replaceNB: info[9] || ' ',
        replaceMB: info[10] || ' ',
        replaceYB: info[11] || ' ',
        replaceSOP: info[12] || ' '
    });
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
    main(path.join(__dirname, "output.docx"));
});


app.post('/submit3', (req, res) => {
    const info = req.body.info;
    const content = fs.readFileSync(
        path.resolve(__dirname, "example/Shablon2.docx"),
        "binary"
    );
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.render({
        replaceProduct: info[0] || ' ',
        replaceFIO: info[1] || ' ',
        replaceIndex: info[2] || ' ',
        replaceCountry: info[3] || ' ',
        replaceCity: info[4] || ' ',
        replaceAdress: info[5] || ' ',
        replaceAdressNum: info[6] || ' ',
        replaceCorpus: info[7] || ' ',
        replaceNumbH: info[8] || ' ',
        replaceSeries: info[9] || ' ',
        replacePNumber: info[10] || ' ',
        replaceDateGive: info[11] || ' ',
        replaceWhos: info[12] || ' ',
        replaceOP: info[13] || ' '
    });
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
    main(path.join(__dirname, "output.docx"));

});
app.post('/submit4', (req, res) => {
    const info = req.body.info;
    const content = fs.readFileSync(
        path.resolve(__dirname, "example/Shablon3.docx"),
        "binary"
    );
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.render({
        replaceProductName: info[0] || ' ',
        replaceFIO: info[1] || ' ',
        rI: info[2] || ' ',
        rC: info[3] || ' ',
        rCi: info[4] || ' ',
        rA: info[5] || ' ',
        rAN: info[6] || ' ',
        rCo: info[7] || ' ',
        rNH: info[8] || ' ',
        replaceTelephone: info[16] || ' ',
        replaceEmail: info[17] || ' ',
        replaceYearReg: info[24] || ' ',
        replaceSeries: info[12] || ' ',
        replacePNumber: info[13] || ' ',
        replaceSnils: info[14] || ' '
    })
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);

    const content2 = fs.readFileSync(
        path.resolve(__dirname, "example/Shablon4.docx"),
        "binary"
    );
    const zip2 = new PizZip(content2);
    const doc2 = new Docxtemplater(zip2, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc2.render({
        ath: info[18] || ' ',
        replaceFIO: info[1] || ' ',
        rI: info[2] || ' ',
        rC: info[3] || ' ',
        rCi: info[4] || ' ',
        rA: info[5] || ' ',
        rAN: info[6] || ' ',
        rCo: info[7] || ' ',
        rNH: info[8] || ' ',
        replaceNB: info[9] || ' ',
        replaceMB: info[10] || ' ',
        replaceYB: info[11] || ' ',
        replaceGR: info[15] || ' ',
        replaceShortLore: info[19] || ' ',

        ids: info[20] || ' ',
        ref: info[21] || ' ',
        Do: info[22] || ' ',
        Dp: info[23] || ' '
    })
    const buf2 = doc2.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output2.docx"), buf2);
    mainforDop(path.join(__dirname, "output.docx"), path.join(__dirname, "output2.docx"));

});
app.post('/submit5', (req, res) => {
    const info = req.body.info;
    var content = fs.readFileSync(
        path.resolve(__dirname, "example/Shablon5.docx"),
        "binary"
    );
    if (info[17] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.1.docx"),
            "binary"
        );
    } else if (info[33] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.2.docx"),
            "binary"
        );
    } else if (info[49] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.3.docx"),
            "binary"
        );
    } else if (info[65] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.4.docx"),
            "binary"
        );
    } else if (info[81] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.5.docx"),
            "binary"
        );
    } else if (info[97] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.6.docx"),
            "binary"
        );
    } else if (info[113] == "") {
        content = fs.readFileSync(
            path.resolve(__dirname, "example/Shablon5.7.docx"),
            "binary"
        );
    }
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    if (info[17] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' '

        });
    } else if (info[33] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' '

        });
    } else if (info[49] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' ',

            replaceFIO3: info[33] || ' ',
            replaceIndex3: info[34] || ' ',
            replaceCountry3: info[35] || ' ',
            replaceCity3: info[36] || ' ',
            replaceAdress3: info[37] || ' ',
            replaceAdressNum3: info[38] || ' ',
            replaceCorpus3: info[39] || ' ',
            replaceNumbH3: info[40] || ' ',
            replaceNB3: info[41] || ' ',
            replaceMB3: info[42] || ' ',
            replaceYB3: info[43] || ' ',
            rSN3: info[44] || ' ',
            rPN3: info[45] || ' ',
            replaceSnils3: info[46] || ' ',
            replaceShortLore3: info[47] || ' ',
            replaceGR3: info[48] || ' '

        });
    } else if (info[65] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' ',

            replaceFIO3: info[33] || ' ',
            replaceIndex3: info[34] || ' ',
            replaceCountry3: info[35] || ' ',
            replaceCity3: info[36] || ' ',
            replaceAdress3: info[37] || ' ',
            replaceAdressNum3: info[38] || ' ',
            replaceCorpus3: info[39] || ' ',
            replaceNumbH3: info[40] || ' ',
            replaceNB3: info[41] || ' ',
            replaceMB3: info[42] || ' ',
            replaceYB3: info[43] || ' ',
            rSN3: info[44] || ' ',
            rPN3: info[45] || ' ',
            replaceSnils3: info[46] || ' ',
            replaceShortLore3: info[47] || ' ',
            replaceGR3: info[48] || ' ',

            replaceFIO4: info[49] || ' ',
            replaceIndex4: info[50] || ' ',
            replaceCountry4: info[51] || ' ',
            replaceCity4: info[52] || ' ',
            replaceAdress4: info[53] || ' ',
            replaceAdressNum4: info[54] || ' ',
            replaceCorpus4: info[55] || ' ',
            replaceNumbH4: info[56] || ' ',
            replaceNB4: info[57] || ' ',
            replaceMB4: info[58] || ' ',
            replaceYB4: info[59] || ' ',
            rSN4: info[60] || ' ',
            rPN4: info[61] || ' ',
            replaceSnils4: info[62] || ' ',
            replaceShortLore4: info[63] || ' ',
            replaceGR4: info[64] || ' '
        });
    } else if (info[81] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' ',

            replaceFIO3: info[33] || ' ',
            replaceIndex3: info[34] || ' ',
            replaceCountry3: info[35] || ' ',
            replaceCity3: info[36] || ' ',
            replaceAdress3: info[37] || ' ',
            replaceAdressNum3: info[38] || ' ',
            replaceCorpus3: info[39] || ' ',
            replaceNumbH3: info[40] || ' ',
            replaceNB3: info[41] || ' ',
            replaceMB3: info[42] || ' ',
            replaceYB3: info[43] || ' ',
            rSN3: info[44] || ' ',
            rPN3: info[45] || ' ',
            replaceSnils3: info[46] || ' ',
            replaceShortLore3: info[47] || ' ',
            replaceGR3: info[48] || ' ',

            replaceFIO4: info[49] || ' ',
            replaceIndex4: info[50] || ' ',
            replaceCountry4: info[51] || ' ',
            replaceCity4: info[52] || ' ',
            replaceAdress4: info[53] || ' ',
            replaceAdressNum4: info[54] || ' ',
            replaceCorpus4: info[55] || ' ',
            replaceNumbH4: info[56] || ' ',
            replaceNB4: info[57] || ' ',
            replaceMB4: info[58] || ' ',
            replaceYB4: info[59] || ' ',
            rSN4: info[60] || ' ',
            rPN4: info[61] || ' ',
            replaceSnils4: info[62] || ' ',
            replaceShortLore4: info[63] || ' ',
            replaceGR4: info[64] || ' ',

            replaceFIO5: info[65] || ' ',
            replaceIndex5: info[66] || ' ',
            replaceCountry5: info[67] || ' ',
            replaceCity5: info[68] || ' ',
            replaceAdress5: info[69] || ' ',
            replaceAdressNum5: info[70] || ' ',
            replaceCorpus5: info[71] || ' ',
            replaceNumbH5: info[72] || ' ',
            replaceNB5: info[73] || ' ',
            replaceMB5: info[74] || ' ',
            replaceYB5: info[75] || ' ',
            rSN5: info[76] || ' ',
            rPN5: info[77] || ' ',
            replaceSnils5: info[78] || ' ',
            replaceShortLore5: info[79] || ' ',
            replaceGR5: info[80] || ' '
        });
    } else if (info[97] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' ',

            replaceFIO3: info[33] || ' ',
            replaceIndex3: info[34] || ' ',
            replaceCountry3: info[35] || ' ',
            replaceCity3: info[36] || ' ',
            replaceAdress3: info[37] || ' ',
            replaceAdressNum3: info[38] || ' ',
            replaceCorpus3: info[39] || ' ',
            replaceNumbH3: info[40] || ' ',
            replaceNB3: info[41] || ' ',
            replaceMB3: info[42] || ' ',
            replaceYB3: info[43] || ' ',
            rSN3: info[44] || ' ',
            rPN3: info[45] || ' ',
            replaceSnils3: info[46] || ' ',
            replaceShortLore3: info[47] || ' ',
            replaceGR3: info[48] || ' ',

            replaceFIO4: info[49] || ' ',
            replaceIndex4: info[50] || ' ',
            replaceCountry4: info[51] || ' ',
            replaceCity4: info[52] || ' ',
            replaceAdress4: info[53] || ' ',
            replaceAdressNum4: info[54] || ' ',
            replaceCorpus4: info[55] || ' ',
            replaceNumbH4: info[56] || ' ',
            replaceNB4: info[57] || ' ',
            replaceMB4: info[58] || ' ',
            replaceYB4: info[59] || ' ',
            rSN4: info[60] || ' ',
            rPN4: info[61] || ' ',
            replaceSnils4: info[62] || ' ',
            replaceShortLore4: info[63] || ' ',
            replaceGR4: info[64] || ' ',

            replaceFIO5: info[65] || ' ',
            replaceIndex5: info[66] || ' ',
            replaceCountry5: info[67] || ' ',
            replaceCity5: info[68] || ' ',
            replaceAdress5: info[69] || ' ',
            replaceAdressNum5: info[70] || ' ',
            replaceCorpus5: info[71] || ' ',
            replaceNumbH5: info[72] || ' ',
            replaceNB5: info[73] || ' ',
            replaceMB5: info[74] || ' ',
            replaceYB5: info[75] || ' ',
            rSN5: info[76] || ' ',
            rPN5: info[77] || ' ',
            replaceSnils5: info[78] || ' ',
            replaceShortLore5: info[79] || ' ',
            replaceGR5: info[80] || ' ',

            replaceFIO6: info[81] || ' ',
            replaceIndex6: info[82] || ' ',
            replaceCountry6: info[83] || ' ',
            replaceCity6: info[84] || ' ',
            replaceAdress6: info[85] || ' ',
            replaceAdressNum6: info[86] || ' ',
            replaceCorpus6: info[87] || ' ',
            replaceNumbH6: info[88] || ' ',
            replaceNB6: info[89] || ' ',
            replaceMB6: info[90] || ' ',
            replaceYB6: info[91] || ' ',
            rSN6: info[92] || ' ',
            rPN6: info[93] || ' ',
            replaceSnils6: info[94] || ' ',
            replaceShortLore6: info[95] || ' ',
            replaceGR6: info[96] || ' '
        });
    } else if (info[113] == "") {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' ',

            replaceFIO3: info[33] || ' ',
            replaceIndex3: info[34] || ' ',
            replaceCountry3: info[35] || ' ',
            replaceCity3: info[36] || ' ',
            replaceAdress3: info[37] || ' ',
            replaceAdressNum3: info[38] || ' ',
            replaceCorpus3: info[39] || ' ',
            replaceNumbH3: info[40] || ' ',
            replaceNB3: info[41] || ' ',
            replaceMB3: info[42] || ' ',
            replaceYB3: info[43] || ' ',
            rSN3: info[44] || ' ',
            rPN3: info[45] || ' ',
            replaceSnils3: info[46] || ' ',
            replaceShortLore3: info[47] || ' ',
            replaceGR3: info[48] || ' ',

            replaceFIO4: info[49] || ' ',
            replaceIndex4: info[50] || ' ',
            replaceCountry4: info[51] || ' ',
            replaceCity4: info[52] || ' ',
            replaceAdress4: info[53] || ' ',
            replaceAdressNum4: info[54] || ' ',
            replaceCorpus4: info[55] || ' ',
            replaceNumbH4: info[56] || ' ',
            replaceNB4: info[57] || ' ',
            replaceMB4: info[58] || ' ',
            replaceYB4: info[59] || ' ',
            rSN4: info[60] || ' ',
            rPN4: info[61] || ' ',
            replaceSnils4: info[62] || ' ',
            replaceShortLore4: info[63] || ' ',
            replaceGR4: info[64] || ' ',

            replaceFIO5: info[65] || ' ',
            replaceIndex5: info[66] || ' ',
            replaceCountry5: info[67] || ' ',
            replaceCity5: info[68] || ' ',
            replaceAdress5: info[69] || ' ',
            replaceAdressNum5: info[70] || ' ',
            replaceCorpus5: info[71] || ' ',
            replaceNumbH5: info[72] || ' ',
            replaceNB5: info[73] || ' ',
            replaceMB5: info[74] || ' ',
            replaceYB5: info[75] || ' ',
            rSN5: info[76] || ' ',
            rPN5: info[77] || ' ',
            replaceSnils5: info[78] || ' ',
            replaceShortLore5: info[79] || ' ',
            replaceGR5: info[80] || ' ',

            replaceFIO6: info[81] || ' ',
            replaceIndex6: info[82] || ' ',
            replaceCountry6: info[83] || ' ',
            replaceCity6: info[84] || ' ',
            replaceAdress6: info[85] || ' ',
            replaceAdressNum6: info[86] || ' ',
            replaceCorpus6: info[87] || ' ',
            replaceNumbH6: info[88] || ' ',
            replaceNB6: info[89] || ' ',
            replaceMB6: info[90] || ' ',
            replaceYB6: info[91] || ' ',
            rSN6: info[92] || ' ',
            rPN6: info[93] || ' ',
            replaceSnils6: info[94] || ' ',
            replaceShortLore6: info[95] || ' ',
            replaceGR6: info[96] || ' ',

            replaceFIO7: info[97] || ' ',
            replaceIndex7: info[98] || ' ',
            replaceCountry7: info[99] || ' ',
            replaceCity7: info[100] || ' ',
            replaceAdress7: info[101] || ' ',
            replaceAdressNum7: info[102] || ' ',
            replaceCorpus7: info[103] || ' ',
            replaceNumbH7: info[104] || ' ',
            replaceNB7: info[105] || ' ',
            replaceMB7: info[106] || ' ',
            replaceYB7: info[107] || ' ',
            rSN7: info[108] || ' ',
            rPN7: info[109] || ' ',
            replaceSnils7: info[110] || ' ',
            replaceShortLore7: info[111] || ' ',
            replaceGR7: info[112] || ' '
        });
    } else {
        doc.render({
            replaceProductName: info[0] || ' ',
            replaceFIO1: info[1] || ' ',
            replaceIndex1: info[2] || ' ',
            replaceCountry1: info[3] || ' ',
            replaceCity1: info[4] || ' ',
            replaceAdress1: info[5] || ' ',
            replaceAdressNum1: info[6] || ' ',
            replaceCorpus1: info[7] || ' ',
            replaceNumbH1: info[8] || ' ',
            replaceNB1: info[9] || ' ',
            replaceMB1: info[10] || ' ',
            replaceYB1: info[11] || ' ',
            rSN1: info[12] || ' ',
            rPN1: info[13] || ' ',
            replaceSnils1: info[14] || ' ',
            replaceShortLore1: info[15] || ' ',
            replaceGR1: info[16] || ' ',

            replaceFIO2: info[17] || ' ',
            replaceIndex2: info[18] || ' ',
            replaceCountry2: info[19] || ' ',
            replaceCity2: info[20] || ' ',
            replaceAdress2: info[21] || ' ',
            replaceAdressNum2: info[22] || ' ',
            replaceCorpus2: info[23] || ' ',
            replaceNumbH2: info[24] || ' ',
            replaceNB2: info[25] || ' ',
            replaceMB2: info[26] || ' ',
            replaceYB2: info[27] || ' ',
            rSN2: info[28] || ' ',
            rPN2: info[29] || ' ',
            replaceSnils2: info[30] || ' ',
            replaceShortLore2: info[31] || ' ',
            replaceGR2: info[32] || ' ',

            replaceFIO3: info[33] || ' ',
            replaceIndex3: info[34] || ' ',
            replaceCountry3: info[35] || ' ',
            replaceCity3: info[36] || ' ',
            replaceAdress3: info[37] || ' ',
            replaceAdressNum3: info[38] || ' ',
            replaceCorpus3: info[39] || ' ',
            replaceNumbH3: info[40] || ' ',
            replaceNB3: info[41] || ' ',
            replaceMB3: info[42] || ' ',
            replaceYB3: info[43] || ' ',
            rSN3: info[44] || ' ',
            rPN3: info[45] || ' ',
            replaceSnils3: info[46] || ' ',
            replaceShortLore3: info[47] || ' ',
            replaceGR3: info[48] || ' ',

            replaceFIO4: info[49] || ' ',
            replaceIndex4: info[50] || ' ',
            replaceCountry4: info[51] || ' ',
            replaceCity4: info[52] || ' ',
            replaceAdress4: info[53] || ' ',
            replaceAdressNum4: info[54] || ' ',
            replaceCorpus4: info[55] || ' ',
            replaceNumbH4: info[56] || ' ',
            replaceNB4: info[57] || ' ',
            replaceMB4: info[58] || ' ',
            replaceYB4: info[59] || ' ',
            rSN4: info[60] || ' ',
            rPN4: info[61] || ' ',
            replaceSnils4: info[62] || ' ',
            replaceShortLore4: info[63] || ' ',
            replaceGR4: info[64] || ' ',

            replaceFIO5: info[65] || ' ',
            replaceIndex5: info[66] || ' ',
            replaceCountry5: info[67] || ' ',
            replaceCity5: info[68] || ' ',
            replaceAdress5: info[69] || ' ',
            replaceAdressNum5: info[70] || ' ',
            replaceCorpus5: info[71] || ' ',
            replaceNumbH5: info[72] || ' ',
            replaceNB5: info[73] || ' ',
            replaceMB5: info[74] || ' ',
            replaceYB5: info[75] || ' ',
            rSN5: info[76] || ' ',
            rPN5: info[77] || ' ',
            replaceSnils5: info[78] || ' ',
            replaceShortLore5: info[79] || ' ',
            replaceGR5: info[80] || ' ',

            replaceFIO6: info[81] || ' ',
            replaceIndex6: info[82] || ' ',
            replaceCountry6: info[83] || ' ',
            replaceCity6: info[84] || ' ',
            replaceAdress6: info[85] || ' ',
            replaceAdressNum6: info[86] || ' ',
            replaceCorpus6: info[87] || ' ',
            replaceNumbH6: info[88] || ' ',
            replaceNB6: info[89] || ' ',
            replaceMB6: info[90] || ' ',
            replaceYB6: info[91] || ' ',
            rSN6: info[92] || ' ',
            rPN6: info[93] || ' ',
            replaceSnils6: info[94] || ' ',
            replaceShortLore6: info[95] || ' ',
            replaceGR6: info[96] || ' ',

            replaceFIO7: info[97] || ' ',
            replaceIndex7: info[98] || ' ',
            replaceCountry7: info[99] || ' ',
            replaceCity7: info[100] || ' ',
            replaceAdress7: info[101] || ' ',
            replaceAdressNum7: info[102] || ' ',
            replaceCorpus7: info[103] || ' ',
            replaceNumbH7: info[104] || ' ',
            replaceNB7: info[105] || ' ',
            replaceMB7: info[106] || ' ',
            replaceYB7: info[107] || ' ',
            rSN7: info[108] || ' ',
            rPN7: info[109] || ' ',
            replaceSnils7: info[110] || ' ',
            replaceShortLore7: info[111] || ' ',
            replaceGR7: info[112] || ' ',

            replaceFIO8: info[113] || ' ',
            replaceIndex8: info[114] || ' ',
            replaceCountry8: info[115] || ' ',
            replaceCity8: info[116] || ' ',
            replaceAdress8: info[117] || ' ',
            replaceAdressNum8: info[118] || ' ',
            replaceCorpus8: info[119] || ' ',
            replaceNumbH8: info[120] || ' ',
            replaceNB8: info[121] || ' ',
            replaceMB8: info[122] || ' ',
            replaceYB8: info[123] || ' ',
            rSN8: info[124] || ' ',
            rPN8: info[125] || ' ',
            replaceSnils8: info[126] || ' ',
            replaceShortLore8: info[127] || ' ',
            replaceGR8: info[128] || ' '
        });
    }
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
    main(path.join(__dirname, "output.docx"));
});

app.post('/submit6', (req, res) => {
    const info = req.body.info;
    const content = fs.readFileSync(
        path.resolve(__dirname, "example/CODT.docx"),
        "binary"
    );
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.render({
        replaceProductName: info[1] || ' ',
        replaceAuthors: info[0] || ' ',
        replaceCode: info[2] || ' '
    });
    const buf = doc.getZip().generate({
        type: "nodebuffer",
        // compression: DEFLATE adds a compression step.
        // For a 50MB output document, expect 500ms additional CPU time
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
    main(path.join(__dirname, "output.docx"));

});


app.listen(4000, () => {
    console.log('App is running...');
    const open = require('open');
    open(`http://localhost:4000`);
});
