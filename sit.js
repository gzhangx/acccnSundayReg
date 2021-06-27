
const isLocal = true;

const gs = require('./getSheet');
const fs = require('fs');
const request = require('superagent');
const { get, sortBy } = require('lodash');


const sheetName = 'Sheet1';
const credentials = require('./credentials.json')
async function myFunction() {
const client = await gs.getClient('gzprem');
const sheet = client.getSheetOps(credentials.sheetId);
const fixedInfo = await sheet.read(`'Fixed'!A1:D30`);

//const preSits = credentials.preSits || [];
const preSits = fixedInfo.values.map(f => {
  if (f[2])
    return `${f[0].trim()}:${f[1].trim()}`;
  return null;
}).filter(x => x);


const getDisplayRow = r => r + 1; //1 based
function parseSits() {
  const lines = fs.readFileSync('./sitConfig.txt').toString().split('\n');
  const starts = lines[0].split('\t').reduce((acc, l,i) => {
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
        if (!blk.rowColMin[curRow] && blk.rowColMin[curRow]!== 0 ) blk.rowColMin[curRow] = i;
        if (i <= (blk.rowColMin[curRow] || 0)) blk.rowColMin[curRow] = i;
        if (i >= (blk.rowColMax[curRow] || 0)) blk.rowColMax[curRow] = i;
        blk.maxRow = curRow;
        blk.sits.push({
          col: i,
          row: curRow,
        })
      }
      return acc;
    },acc)
  }, []).map(b => {    
    return {
      letterCol: b.sits[0].col === b.min ? 0:b.max - b.min,
      ...b,
      cols: b.max - b.min + 1,
      rows: b.maxRow - b.minRow + 1,
      sits: b.sits.map(s => {
        const rowColMin = b.rowColMin[s.row];
        const rowColMax = b.rowColMax[s.row];
        const rowCols = rowColMax - rowColMin;
        const colPos = s.col - rowColMin;
        return ({
          side: colPos < rowCols/3?'A': colPos> rowCols*2/3?'C':'B',
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
  return blkInfo.map((b,bi) => {
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
      goodRowsToUse: rows.map((r,i) => {
        return !(i%2) || i === rows.length-1
      }),
      sits: rows,
    };
  });
}
const pureSitConfig = parseSits();
//console.log(pureSitConfig.map(s=>({cols: s.cols, rows: s.rows})))
//return console.log(pureSitConfig.map(r => r.sits.map(v => v.map(vv => vv ? 'X' : ' ').join('')).join('\n')).join('\n'));

function getNextSundays() {
  let cur = new Date();
  const oneday = 24 * 60 * 60 * 1000;
  while (cur.getDay() !== 0) {    
    cur = new Date(cur.getTime() + oneday);    
  }
  const res = [];
  for (let i = 0; i < 10; i++) {
    res[i] =  (getDateStr(new Date(cur.getTime()+(oneday*i))));
  }
  return res;
}

function getDateStr(date) {
  return `${date.getFullYear()}-${(date.getMonth()+1).toString().padStart(2,'0')}-${date.getDate().toString().padStart(2,'0')}`;
}


  const nextSundays = getNextSundays();
  const nextSunday = nextSundays[0];
  console.log(`nextSunday=${nextSunday}`);
  const authorizationToken = credentials.eventBriteAuth;
  const ebFetch = async url => {
    console.log(`url=${url}`);
    if (isLocal) {
      //var response = await request.get("https://www.eventbriteapi.com/v3/events/156798329023/attendees").set('Authorization', authorizationToken).send();
      //pages = response.body
      const response = await require('superagent').get(url).set('Authorization', authorizationToken).send();
      return response.body;
    } else {
      var response = UrlFetchApp.fetch(url, {
        headers:
        {
          authorization: authorizationToken
        }
      });
      const pages = JSON.parse(response.getContentText())
      return pages;
    }
  }
  const eventArys = await ebFetch('https://www.eventbriteapi.com/v3/organizations/544694808143/events/?name_filter=' + encodeURIComponent(credentials.eventTitle) + '&time_filter=current_future');
  const eventsMapped = eventArys.events.map(e => {
    return {
      id: e.id,
      date: e.start.local.slice(0, 10),
      name: e.name,
      status: e.status,
    }
  }).filter(s=>s.status === 'live');
  console.log(eventsMapped)
  let nextGoodEvent = (eventsMapped.filter(x => x.date === nextSunday))[0];
  let nsi = 0;
  while (!nextGoodEvent && nsi < nextSundays.length) {
    nsi++;
    nextGoodEvent = (eventsMapped.filter(x => x.date === nextSundays[nsi]))[0];
  }
  if (!nextGoodEvent) {
    console.log('Next event not found');
    console.log(eventsMapped);
    return;
  }
  
  let pages = await ebFetch(`https://www.eventbriteapi.com/v3/events/${nextGoodEvent.id}/attendees`);
  if (false && isLocal) {    
    const fakeSizes = { 'gg1': 3, 'gg12': 4 ,'gg19':6,'gg22':8,'gg33':5};
    const fakeNames = [];
    for (let i = 0; i < 120; i++) {
      fakeNames[i] = 'gg' + i;
    }
    pages = {
      attendees:
        pages = fakeNames.map(n => {
          return {
            quantity: fakeSizes[n] || 1,
            profile: {
              name: n,
              email: n + '@hotmail.com'
            }
          }
        })
    }
  }

  const preSiteItems = fixedInfo.values.filter(v => v[2]).map((v,pos) => {    
    const name = v[0];
    const email = v[1];
    const blkRowId = v[2];
    const rc = blkRowId.slice(1).split('-');
    const key = `${name}:${email}`.toLocaleLowerCase();
    return {
      quantity: 1,
      emails: [email],
      names: [name],
      name,
      key,
      pos,
      id: pos + 1,
      blkRowId,
      posInfo: {
        block: blkRowId[0],
        row: parseInt(rc[0]),
        rowInfo: null,
        side: rc[1],
        //block: blkMap[blki],
        //row,
        //rowInfo: curRow[0],
        //side: 'A',
      }
    }
  });
  const preSiteItemsByBlkRowId = preSiteItems.reduce((acc, r) => {
    acc[r.blkRowId] = r;
    return acc;
  }, {});
  //return console.log(preSiteItems)
  const names = sortBy(pages.attendees.reduce((acc, att) => {
    if (att.cancelled) return acc;
    let ord = acc.oid[att.order_id];
    const key = `${att.profile.name}:${att.profile.email}`.toLocaleLowerCase();
    //console.log(`attend ${key} order ${att.order_id} ${att.cancelled}  ${att.status}`);
    const existing = preSits.find(k => k.toLowerCase() === key);
    if (existing) {
      return acc;
    }
    if (!ord) {
      ord = {
        quantity: 0,
        emails: [],
        names: [],
        key,
        pos: acc.ary.length,
        id: acc.ary.length + 1,
        name: att.profile.name,
      };
      acc.oid[att.order_id] = ord;      
      acc.ary.push(ord);
    }
    ord.quantity++;
    ord.emails.push(att.profile.email);
    ord.names.push(att.profile.name);
    return acc;
  }, {
    ary: preSiteItems, oid: {}
  }).ary,'name');

  let colors = [[0, 0, 255], [0, 255, 0], [255, 0, 0], [0, 255, 255], [255, 0, 255], [255, 200, 200]];
  let fontColor = ['#ffff00', '#ff00ff', '#00ffff', '#000000', '#000000', '#000000'];
  let rgbFontColor = [[255, 255, 0], [255, 0, 255], [0, 255, 255], [0, 0, 0], [0, 0, 0], [0, 0, 0]]
  while (colors.length < names.length) {
    colors.map(c => c.map(c => parseInt(c / 2))).forEach(c => {
      if (c[0] + c[1] + c[2] < 255 + 128) {
        fontColor.push('#ffffff');
        rgbFontColor.push([255,255,255])
      } else {
        fontColor.push('#000000')
        rgbFontColor.push([0, 0, 0])
      }
      return colors.push(c);
    });
  }

  const rgb255toClr = rgb => ['red', 'green', 'blue'].reduce((acc, name, i) => {
    acc[name] = rgb[i] / 255.0;
    return acc;
  }, {});
  const rgbColors = colors.map(rgb255toClr);
  rgbFontColor = rgbFontColor.map(rgb255toClr)
  colors = colors.map(c => `#${c.map(c => c.toString(16).padStart(2,'0')).join('')}`);


  const blockSpacing = 2;
  const fMax = (acc, cr) => acc < cr ? cr : acc;
  //const blockColMaxes = blockConfig.map(r => r.reduce(fMax, 0));
  const blockColMaxes = pureSitConfig.map(r=>r.cols);
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


  const blkMap = ['A', 'B', 'C', 'D']
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

  //const headers = [];
  //for (let i = 0; i < numCols + STARTCol; i++) headers[i] = '';
  /*
  blockStarts.forEach((bs,i)=>{
    const headers = [];
    for (let j = 0; j < blockColMaxes[i];j++) headers[j] = '';  
    sheet.getRange(STARTRow, blockStarts[i], 1, blockColMaxes[i]).setValues([headers]);
    sheet.setColumnWidths(blockStarts[i], blockColMaxes[i], CELLSIZE);
    sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
    sheet.getRange(STARTRow,blockStarts[i], numRows, blockColMaxes[i]).setBackground('yellow');
  });
  */
  
  
  //sheet.getRange(1,1,1,numCols + STARTCol).setValues([headers]);
  //sheet.setColumnWidths(STARTCol, numCols, CELLSIZE);
  //sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
  //sheet.getRange(STARTRow,STARTCol, numRows, numCols).setBackground('yellow');



  const siteSpacing = 3;
  const fitChore = who => {
    if (who.posInfo) return true;
    const blki = 1; //block B only
    const curBlock = blockSits[blki];
    for (let row = 0; row < numRows; row++) {
      const curRow = curBlock[row]?.filter(x => x);
      if (!curRow) break;
      for (let i = 0; i < curRow.length; i++) {
        const cri = curRow[i];
        if (!cri) continue;
        if (cri.user) continue;
        cri.user = who;
        who.posInfo = {
          block: blkMap[blki],
          row,
          rowInfo: cri,
          side: cri.side,
        }
        return true;
      }
    }
    return false;
  }
  const fit = (who, reverse = false) => {
    if (who.posInfo) return true;
    let fited = false;
    for (let rowInc = 0; rowInc < numRows; rowInc++) {      
      const row = reverse ? numRows - rowInc - 1 : rowInc;
      for (let blki = 0; blki < blockSits.length; blki++) {
        if (!pureSitConfig[blki].goodRowsToUse[rowInc]) continue;
        if (credentials.ignoreBlocks[blki]) continue;
        const curBlock = blockSits[blki];
        //if (!curBlock) continue;
        const curRow = curBlock[row]?.filter(x=>x);
        if (!curRow) continue;
        ['left', 'right'].forEach(side => {
          if (fited) return;
          if (side === 'left') {
            if (curRow[0].user) return;
            for (let i = 0; i < who.quantity; i++){
              if (!curRow[i].user)
                curRow[i].user = who;
              else return fit(who, reverse); //this needs testing, i.e. we grow out of current row.
            } 
            who.posInfo = {
              block: blkMap[blki],
              row,
              rowInfo: curRow[0],
              side: 'A',              
            }
            fited = true;
            return;
          } else if (side === 'right') {
            const ind = curRow.length - 1;
            if (curRow[ind].user) return;
            const toSearch = ind - who.quantity - siteSpacing;
            for (let i = ind; i >= toSearch; i--) {
              if (curRow[i].user) return;
            }
            for (let i = 0; i < who.quantity; i++) {
              curRow[ind - i].user = who;
            }
            fited = true;
            who.posInfo = {
              block: blkMap[blki],
              row,
              rowInfo: curRow[ind],
              side: 'C',
            }
          }
        });
      }
    }
    if (!fited) {
      let maxAva = 0;
      let curMaxRow = null;
      let curMaxRowNumber = -1;
      let curMaxRowBlk = -1;
      let curMaxRowStart = -1;
      let curMaxRowEnd = -1;
      for (let row = 0; row < numRows; row++) {
        for (let blki = 0; blki < blockSits.length; blki++) {
          if (!pureSitConfig[blki].goodRowsToUse[row]) continue;
          if (credentials.ignoreBlocks[blki]) continue;
          const curBlock = blockSits[blki];
          if (!curBlock) continue;
          if (!curBlock[row]) {
            continue;
          }
          const curRow = curBlock[row].filter(x=>x);
          //let rowTotal = 0;
          let curAvaStart = -1;
          let curAvaLen = -1;
          for (let i = 0; i < curRow.length; i++) {
            if (!curRow[i].user) {
              if (curAvaStart < 0) {
                curAvaStart = i;
                curAvaLen = 1;
              } else {
                curAvaLen = i - curAvaStart + 1;
              }
              if (curAvaLen > maxAva) {
                maxAva = curAvaLen;
                curMaxRow = curRow;
                curMaxRowNumber = row;
                curMaxRowBlk = blki;
                curMaxRowStart = curAvaStart;
                curMaxRowEnd = i;
              }
            } else {
              curAvaStart = -1;
              curAvaLen = 0;
            }
            //if (!curRow[i].user) rowTotal++;
          }
          //if (rowTotal > maxAva) {
          //  maxAva = rowTotal;
          //  curMaxRow = curRow;
          //  curMaxRowNumber = row;
          //  curMaxRowBlk = blki;
          //}
        }
      }      

      if (curMaxRow) {
        const bestSpacing = {
          start: curMaxRowStart,
          end: curMaxRowEnd,
          size: curMaxRowEnd - curMaxRowStart +1,
        };
          
        
        if (bestSpacing) {
          if (bestSpacing.size > who.quantity + (siteSpacing * 2)) {
            const left = Math.round((bestSpacing.size - who.quantity) / 2);
            for (let i = 0; i < who.quantity; i++) {
              const curCell = curMaxRow[bestSpacing.start + left + i];
              curCell.user = who;
              who.posInfo = {
                block: blkMap[curMaxRowBlk],
                row: curMaxRowNumber,
                rowInfo: curCell,
                side: curCell.side,
              }
            }
            return true;
          }
        }
      }
    }
    return fited;
  };

  const choreNames = ['詩 ','詩-']
  names.filter(n => choreNames.find(c => n.name.startsWith(c))).forEach(n => {
    fitChore(n);
  })
  names.forEach(n => {
    if (!fit(n)) {
      console.log(`Warning, unable to fit ${n.name}`)
    }    
  });


  if (!isLocal) {
    const sittingSheet = SpreadsheetApp.openById('1p7W0Gwh88tCSiEA7Y6S_tVfyTMMbzotjaSZmfgEpDHY');

    const sheet = sittingSheet.getSheets()[0];
    sheet.clear();
    sheet.setColumnWidths(STARTCol, numCols, CELLSIZE);
    sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
    blockSits.forEach(blk => {
      blk.forEach(row => {
        row.forEach(c => {
          const range = sheet.getRange(c.uiPos.row, c.uiPos.col);
          range.setValue(c.user ? c.user.id : '-');
          if (c.user) {
            range.setBackground(colors[c.user.id]);
            range.setFontColor(fontColor[c.user.id]);
          } else {
            range.setBackground('yellow');
          }
        });
      })
    });
    const userInfo = [
      ['Code', 'Name', 'Quantity', 'Email'],
      ...names.map(n => [n.id, n.names.join(','), n.quantity, n.emails.join(',')])
    ];
    sheet.getRange(namesStartRow, 1, names.length + 1, 4).setValues(
      userInfo
    )

  } else {
    
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
    pureSitConfig.forEach((bc,i) => {
      data[STARTRow - 2][bc.letterCol + blockStarts[i]-1] = {
        user: {
        id:blkMap[i]
      }}
    });
    for (let i = 0; i < numRows; i++) {
      data[i+STARTRow-1][0] = {
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

    
    const endColumnIndex = STARTCol + numCols;
    console.log(`end col num=${numCols} ${STARTCol} end=${endColumnIndex}`);
    const sheetInfo = await sheet.sheetInfo(sheetName);
    if (!sheetInfo) {
      console.log(`sheet ${sheetName} not found`);
    }
    const { sheetId } = sheetInfo;
    const userInfo = [
      ['Code', 'Quantity', '','Pos','','', 'Name', 'Email','ActualPos'],
      ...names.filter(f=>f.posInfo).map(n => [n.id, n.quantity, n.pos, n.posInfo.block, getDisplayRow(n.posInfo.row).toString(), n.posInfo.side,  n.names.join(','), n.emails.join(','), `r=${n.posInfo.rowInfo.row} c=${n.posInfo.rowInfo.col}`])
    ];
 
    // console.log('names==>')
    // console.log(names.map(n => {
    //   return {
    //     rowInfof: n.posInfo.rowInfo,
    //     ...n,
    //   }
    // }));
    const userData = userInfo.map(u => {
      return [u[0].toString(), '', '', '', u[1].toString(), '', '', '', '', { type: 'userColor', val: u[2] }, u[3], u[4], u[5], '', u[6], '', '', '', '', '', '', u[7], '', '', '', '', '', '', '', '', '', '',u[8]];
    }).map(r => {
      return {
        values: r.map(o => {
          let stringValue = o;
          if (typeof (o) === 'object') {
            stringValue = '';
          }
          const horizontalAlignment = 'LEFT';
          const cell = {
            userEnteredValue: { stringValue }
          };
          if (typeof (o) === 'object') {
            cell.userEnteredFormat = {
              backgroundColor: rgbColors[o.val],
            }
          }
          return cell;
        })
      };
    });
    const rowData = data.map(r => {
      return {
        values: r.map((cval) => {
          const user = cval && cval.user;
          const stringValue = (cval ? (user?.id || '-') : '').toString();
          const horizontalAlignment = 'CENTER';
          const cell = {
            userEnteredValue: { stringValue }
          };
          if (user && user.id) {
            cell.userEnteredFormat = {
              backgroundColor: rgbColors[user.pos],
              horizontalAlignment,
              textFormat: {
                foregroundColor: rgbFontColor[user.pos],
                //fontFamily: string,
                //"fontSize": integer,
                bold: true,
                //"italic": boolean,
                //"strikethrough": boolean,
                //"underline": boolean,                                            
              },
              borders: {
                bottom: {
                  style: 'SOLID',
                  width: 1,
                  color: {
                    blue: 0,
                    green: 1,
                    red: 0
                  }
                }
              }
            };
          } else {
            cell.userEnteredFormat = {
              horizontalAlignment,
              backgroundColor: cval ? {
                blue: 0,
                green: 1,
                red: 1
              } : {
                blue: 1,
                green: 1,
                red: 1
              },
            }
          }
          return cell;
        })
      };
    });
    
    const endRowIndex = data.length + userData.length + 1;
    const updateData = {
      requests: [
        {
          updateCells: {
            fields: '*',
            range: {
              sheetId,
              startColumnIndex: 0,
              endColumnIndex,
              startRowIndex: 0,
              endRowIndex
            },
            rows: [...rowData, ...userData]
          }
        }
      ]
    };

    if (endRowIndex > sheetInfo.rowCount || endColumnIndex > sheetInfo.columnCount) {
      const requests = [];
      if (endColumnIndex > sheetInfo.columnCount) {
        requests.push({
          appendDimension: {
            sheetId,
            dimension: 'COLUMNS',
            length: endColumnIndex - sheetInfo.columnCount,
          }
        })
      }
      if (endRowIndex > sheetInfo.rowCount) {
        requests.push({
          appendDimension: {
            sheetId,
            dimension: 'ROWS',
            length: endRowIndex - sheetInfo.rowCount ,
          }
        })
      }
      if (requests.length) {        
        console.log(`updating column endColumnIndex=${endColumnIndex} sheetInfo.columnCount=${sheetInfo.columnCount} ${endColumnIndex > sheetInfo.columnCount}` );
        console.log({
          sheetId,
          dimension: 'COLUMNS',
          length: endColumnIndex - sheetInfo.columnCount,
        })
        await sheet.doBatchUpdate({ requests });
        console.log('column updated');
      }
    }
    await sheet.doBatchUpdate({
      requests: [
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'COLUMNS',
              startIndex: 0,
              endIndex: endColumnIndex
            },
            properties: {
              pixelSize: CELLSIZE
            },
            fields: 'pixelSize'
          }
        },
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: 0,
              endIndex: endRowIndex
            },
            properties: {
              pixelSize: CELLSIZE
            },
            fields: 'pixelSize'
          }
        }
      ]
    })
    
    
    //fs.writeFileSync('test.json', JSON.stringify(updateData, null, 2))
    //fs.writeFileSync('debugblockSits.json', JSON.stringify(blockSits,null,2))
    //endRowIndex > sheetInfo.rowCount ? endRowIndex : sheetInfo.rowCount,
    if (sheetInfo.rowCount > endRowIndex) {
      await sheet.doBatchUpdate({
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId,
                dimension: 'ROWS',
                startIndex: endRowIndex,
                endIndex: sheetInfo.rowCount
              }
            }
          },
        ],
      });
    }
    await sheet.doBatchUpdate(updateData);
  }

}

if (isLocal) {
  myFunction().catch(err => {
    console.log(get(err, 'response.body') || err);    
  });
}