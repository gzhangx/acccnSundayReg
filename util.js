const fs = require('fs');
const jimp = require('jimp');
const Promise = require('bluebird');

const gs = require('./getSheet');
const credentials = require('./credentials.json');
const nodemailer = require('nodemailer');

async function getTemplates(sheet) {
    const templates = (await sheet.readValues(`'Template'!A1:C100`));
    return templates;
}
function parseSits(pack=2) {
    const lines = fs.readFileSync('./sitConfig.txt').toString().split('\n');
    const starts = lines[0].split('\t').reduce((acc, l, i) => {
        if (l === 'R') acc.push(i);
        return acc;
    }, []);
    const getBlk = p => {
        if (p > starts[2]) {
            if (p > starts[3]) return 3;
            return 2;
        }
        if (p < starts[1]) return 0;
        return 1;
    };
    const blkInfo = lines.slice(1).reduce((acc, l, curRow) => {
        return l.split('\t').reduce((acc, v, i) => {
            const blki = getBlk(i);
            if (v === 'X' || v === 'N') {
                let blk = acc[blki];
                if (!blk) {
                    blk = { min: i, max: i, minRow: curRow, maxRow: curRow, sits: [], rowColMin: {}, rowColMax: {} };
                    acc[blki] = blk;
                }
                if (blk.min > i) blk.min = i;
                if (blk.max < i) blk.max = i;
                if (!blk.rowColMin[curRow] && blk.rowColMin[curRow] !== 0) blk.rowColMin[curRow] = i;
                if (i <= (blk.rowColMin[curRow] || 0)) blk.rowColMin[curRow] = i;
                if (i >= (blk.rowColMax[curRow] || 0)) blk.rowColMax[curRow] = i;
                blk.maxRow = curRow;
                blk.sits.push({
                    col: i,
                    row: curRow,
                    sitTag: v,
                })
            }
            return acc;
        }, acc)
    }, []).map(b => {
        return {
            letterCol: b.sits[0].col === b.min ? 0 : b.max - b.min,
            ...b,
            cols: b.max - b.min + 1,
            rows: b.maxRow - b.minRow + 1,
            sits: b.sits.map(s => {
                const rowColMin = b.rowColMin[s.row];
                const rowColMax = b.rowColMax[s.row];
                const rowCols = rowColMax - rowColMin;
                const colPos = s.col - rowColMin;
                return ({
                    sitTag: s.sitTag,
                    side: colPos < rowCols / 3 ? 'A' : colPos > rowCols * 2 / 3 ? 'C' : 'B',
                    col: s.col - b.min,
                    row: s.row - b.minRow,
                });
            })
        }
    });
    //console.log(starts);
    //console.log(blkInfo.map(b => ({
    //  cols: b.cols,
    //  rows: b.rows,
    //})));
    return blkInfo.map((b, bi) => {
        const rows = [];
        for (let r = 0; r < b.rows; r++) {
            const rr = [];
            rows[r] = rr;
            for (let c = 0; c < b.cols; c++) {
                rr[c] = null;
            }
        }

        b.sits.forEach(s => {
            rows[s.row][s.col] = {
                ...s,
            };
        })
        //console.log(rows.map(r => r.join('')).join('\n'));
        return {
            ...b,
            goodRowsToUse: rows.map((r, i) => {
                return !(i % pack) || i === rows.length - 1
            }),
            sits: rows,
        };
    });
}

const getDisplayRow = r => r + 1; //1 based

async function initAll() {
    const client = await gs.getClient('gzprem');
    const sheet = client.getSheetOps(credentials.sheetId);


    const templates = await getTemplates(sheet);
    const pack = templates.filter(f => f[0] === 'pack').map(f => parseInt(f[1] || 1))[0] || 2;
    const nextSundays = getNextSundays();
    const nextSunday = nextSundays[0];
    console.log(`nextSunday=${nextSunday}`);

    const fixedInfo = await sheet.readValues(`'${nextSunday}'!A1:F300`).catch(err => {
        console.log('Unable to load fixed')
        console.log(err.response.body);
        return [];
    });

    const blockSpacing = 2;
    const fMax = (acc, cr) => acc < cr ? cr : acc;
    //const blockColMaxes = blockConfig.map(r => r.reduce(fMax, 0));
    const pureSitConfig = parseSits(pack);
    const blockColMaxes = pureSitConfig.map(r => r.cols);
    const numCols = blockColMaxes.reduce((acc, r) => acc + r + blockSpacing, 0);
    //const numRows = blockConfig.map(r => r.length).reduce(fMax, 0);
    const numRows = pureSitConfig.map(r => r.rows).reduce(fMax, 0);

    const STARTCol = 4;
    const STARTRow = 3;
    const namesSpacking = 3;

    const namesStartRow = STARTRow + numRows + namesSpacking;
    const CELLSIZE = 20;
    const blockStarts = blockColMaxes.reduce((acc, b) => {
        const curStart = acc.cur + blockSpacing + acc.prev;
        acc.prev = b;
        acc.res.push(curStart);
        acc.cur = curStart;
        return acc;
    }, {
        res: [],
        prev: 0,
        cur: STARTCol - blockSpacing,
    }).res;


    const blkMap = Object.freeze(['A', 'B', 'C', 'D']);
    const blkLetterToId = blkMap.reduce((acc, ltr, id) => {
        acc[ltr] = id;
        return acc;
    }, {});


    function getDisplayData(blockSits) {
        const data = [];
        const debugCOLLimit = 30;
        for (let i = 0; i < STARTRow + numRows; i++) {
            data[i] = [];
            for (let j = 0; j < STARTCol + numCols; j++) {
                data[i][j] = null;
            }

            //debug
            //data[i] = [];
            for (let j = 0; j < debugCOLLimit; j++)
                data[i][j] = null;
        }

        //top col cord
        pureSitConfig.forEach((bc, i) => {
            data[STARTRow - 2][bc.letterCol + blockStarts[i] - 1] = {
                user: {
                    id: blkMap[i]
                }
            }
        });
        for (let i = 0; i < numRows; i++) {
            data[i + STARTRow - 1][0] = {
                user: {
                    id: getDisplayRow(i).toString()
                }
            }
        }


        blockSits.forEach(blk => {
            blk.forEach(r => {
                r.forEach(c => {
                    if (!c) return;
                    //data[c.uiPos.row - STARTRow][c.uiPos.col - STARTCol] = c.user ? c.user.id : 'e';
                    //if (c.uiPos.col < debugCOLLimit) //debug
                    try {
                        data[c.uiPos.row - 1][c.uiPos.col - 1] = c;
                    } catch (err) {
                        data[c.uiPos.row - 1][c.uiPos.col - 1] = c;
                        throw err;
                    }
                })
            })
        });
        return data;
    }


    function generateBlockSits(preSiteItemsByBlkRowId) {
        const blockSits = pureSitConfig.map((blk, bi) => {
            return blk.sits.map(s => {
                return s.map(r => {
                    if (!r) return null;

                    const blk = {
                        ...r,
                        blkRow: `${blkMap[bi]}${r.row}`,
                        blkRowId: `${blkMap[bi]}${r.row}-${r.col}`,
                        user: null,
                        uiPos: {
                            col: blockStarts[bi] + r.col,
                            row: STARTRow + r.row,
                        }
                    };
                    const user = preSiteItemsByBlkRowId[blk.blkRowId];
                    if (user) {
                        blk.user = user;
                        user.posInfo.rowInfo = blk;
                        //user.posInfo.side = `${blk.side}-${user.posInfo.side}`;
                        user.posInfo.side = blk.side;
                    }
                    return blk;
                });
            });
        });
        return blockSits;
    }


    async function generateImag(key) { //B2-5
        const preSiteItemsByBlkRowId = {
            [key]: {
                id: 'U',
                posInfo: {}
            }
        };
        const blockSits = generateBlockSits(preSiteItemsByBlkRowId);
        const data = getDisplayData(blockSits);
        
        const imgRes = await jimp.loadFont(jimp.FONT_SANS_16_BLACK).then(font => {
            return new Promise((resolve, reject) => {
                new jimp(data[0].length * CELLSIZE, data.length * CELLSIZE, 0xffffffff, (err, image) => {
                    if (err) {
                        console.log(err);
                        return reject(err);
                    }

                    data.forEach((rows, rowInd) => {
                        rows.forEach((cell, colInd) => {
                            if (!cell) return;

                            image.scan(colInd * CELLSIZE, rowInd * CELLSIZE, CELLSIZE - 1, CELLSIZE - 1, function (x, y, idx) {
                                //var red = this.bitmap.data[idx + 0];
                                //var green = this.bitmap.data[idx + 1];
                                //var blue = this.bitmap.data[idx + 2];
                                //var alpha = this.bitmap.data[idx + 3];
                                let r = 0xff;
                                let g = 0;
                                let b = 0xff;
                                const user = cell.user;
                                if (!user) g = 0x99;
                                else if (user.id !== 'U') {
                                    r = 0xe0;
                                    g = 0xe0;
                                    b = 0xe0;
                                    if (blkMap.find(m => m === user.id)) {
                                        r = 0xff;
                                        g = 0xff;
                                        b = 0xff;
                                    }
                                } else {
                                    g = 0xff;
                                    b = 0;
                                }
                                this.bitmap.data[idx + 0] = r;
                                this.bitmap.data[idx + 1] = g;
                                this.bitmap.data[idx + 2] = b;
                                this.bitmap.data[idx + 3] = 0xff;
                            });
                            if (cell.user) {
                                image.print(font, colInd * CELLSIZE, rowInd * CELLSIZE, {
                                    text: cell.user.id,
                                    alignmentX: jimp.HORIZONTAL_ALIGN_CENTER,
                                    alignmentY: jimp.VERTICAL_ALIGN_MIDDLE
                                }, CELLSIZE, CELLSIZE);
                            }
                        });
                    });
                    //image.print(font, 10, 10, 'Hello world!');

                    //image.write('test.png')
                    image.getBase64Async(jimp.MIME_PNG).then(rr => resolve(rr));
                });
            });
        });
        return imgRes;
    }

    //generateImag('B2-5').then(r => {
    //    fs.writeFileSync('test.html',`<img src='${r}' />`)
    //});

    function blockKeyIdToSide(blockSits, keyId) {
        const curBlock = blockSits[blkLetterToId(keyId[0])];
        const rowCol = keyId.slice(1).split('-');
        const row = parseInt(rowCol[0]);
        const curRow = curBlock[row];
        const col = parseInt(rowCol[1]);
        const cri = curRow[col];
        return cri.side;
    }

    function getDateStr(date) {
        return `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
    }

    function getNextSundays() {
        let cur = new Date();
        const oneday = 24 * 60 * 60 * 1000;
        while (cur.getDay() !== 0) {
            cur = new Date(cur.getTime() + oneday);
        }
        const res = [];
        for (let i = 0; i < 10; i++) {
            res[i] = (getDateStr(new Date(cur.getTime() + (oneday * i))));
        }
        return res;
    }

    async function sendEmail() {
        //const sheet = client.getSheetOps(credentials.sheetId);        
    
        //const templates = await getTemplates(sheet);        
    
        const generated = await Promise.map(fixedInfo, async inf => {
            const id = inf[0];
            const name = inf[1];
            const email = inf[2];
            const side = inf[3];
            const key = inf[4];
            const finished = inf[5];
            if (finished) return null;
            //inf[4] = 'sent';
            return {
                id, name, email, side, key,
                imgSrc: await generateImag(key),
            }
        }, { concurrency: 5 });
        //fs.writeFileSync('test.html', generated.filter(x => x).map(g => {
        //    return `Hello ${g.name} (${g.email}), your assigned sit is ${g.side}, please show this email to your usher for their convience.  Thank you!
        //      ${g.key}<br><img src='${g.imgSrc}'/> <br><br>`;
        //}).join('\n'));
        const transporter = nodemailer.createTransport({
            host: 'smtp.office365.com',
            secureConnection: false, // TLS requires secureConnection to be false
            port: 587, // port for secure SMTP
            tls: {
                ciphers: 'SSLv3'
            },
            auth: credentials.msauth
        });
    
        const emailTemplate = templates.filter(f => f[0] === 'emailTemplate')[0][1];
        const sent = await Promise.map(generated.filter(x => x), async g => {
            const html = emailTemplate.replace(/{name}/g, g.name).replace(/\{sit\}/g, g.side)
                .replace(/{imgSrc}/g, g.imgSrc).replace(/{email}/g, g.email)
                .replace(/{key}/g, g.key)
            //console.log(html)
            try {
                console.log(`Sending to ${g.email}`);
                await transporter.sendMail({
                    from: credentials.msauth.user,
                    subject: 'Church siting (教会座位)',
                    to: g.email,
                    //subject: 'Nodemailer is unicode friendly ✔',            
                    html,
                });
            } catch (err) {
                console.log(err);
                console.log(`failed email ${g.name} ${g.email}`);
                return null;
            }
            return g.id;
        }, { concurrency: 2 });


        fixedInfo.forEach(g => {
            if (sent.find(k => k == g[0])) {
                g[5] = 'sent';
            }
        })
        await sheet.updateValues(`'${nextSunday}'!A1:F${fixedInfo.length}`, fixedInfo);
    }

    return {
        sheet,
        templates,
        fixedInfo,
        nextSundays,

        //parseSits,
        getDisplayRow,
        generateImag,
        getNextSundays,

        pureSitConfig,
        blockSpacing,
        fMax,
        //const blockColMaxes = blockConfig.map(r => r.reduce(fMax, 0));
        blockColMaxes,
        numCols,
        //const numRows = blockConfig.map(r => r.length).reduce(fMax, 0);
        numRows,

        STARTCol,
        STARTRow,
        namesSpacking,

        namesStartRow,
        CELLSIZE,
        blockStarts,


        blkMap,
        blkLetterToId,

        generateBlockSits,
        getDisplayData,
        blockKeyIdToSide,
        getTemplates,
        sendEmail,
    };
}
module.exports = {
    initAll,
}

