const fs = require('fs');
const jimp = require('jimp');
function parseSits() {
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
            if (v === 'X') {
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
                return !(i % 2) || i === rows.length - 1
            }),
            sits: rows,
        };
    });
}

const getDisplayRow = r => r + 1; //1 based


function generateImag() {
    const pureSitConfig = parseSits();

    const blockSpacing = 2;
    const fMax = (acc, cr) => acc < cr ? cr : acc;
    //const blockColMaxes = blockConfig.map(r => r.reduce(fMax, 0));
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


    const blkMap = ['A', 'B', 'C', 'D'];
    const blkLetterToId = blkMap.reduce((acc, ltr, id) => {
        acc[ltr] = id;
        return acc;
    }, {});
    const preSiteItemsByBlkRowId = [];
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
    jimp.loadFont(jimp.FONT_SANS_16_BLACK).then(font => {
        new jimp(data[0].length * CELLSIZE, data.length * CELLSIZE, 0x001111ff, (err, image) => {
            if (err) {
                console.log(err);
                return;
            }
            let debugdone = 0;
        
            data.forEach((rows, rowInd) => {
                rows.forEach((cell, colInd) => {
                    if (!cell) return;
                    //if (debugdone> 10) return;
                    debugdone++;
                    console.log(`row=${rowInd} col=${colInd}`)
                    
                    image.scan(colInd * CELLSIZE, rowInd * CELLSIZE, CELLSIZE-1, CELLSIZE-1, function (x, y, idx) {
                        //var red = this.bitmap.data[idx + 0];
                        //var green = this.bitmap.data[idx + 1];
                        //var blue = this.bitmap.data[idx + 2];
                        //var alpha = this.bitmap.data[idx + 3];
                        this.bitmap.data[idx + 0] = 0xff;
                        this.bitmap.data[idx + 2] = 0xff;
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
        
            image.write('test.png')
        });
    });
}

//generateImag();

module.exports = {
    parseSits,
    getDisplayRow,
    generateImag,
}

