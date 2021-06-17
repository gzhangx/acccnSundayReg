
const isLocal = true;

const gs = require('./getSheet');
const fs = require('fs');
const request = require('superagent');
const { get } = require('lodash');
//const sheet = gs.createSheet();
//const cardStatementData = await sheet.readSheet('A', 'ChaseCard!A:F')

const sheetName = 'Sheet1';
const credentials = require('./credentials.json')

const preSits = credentials.preSits;

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
          blk = { min: i, max: i, minRow: curRow, maxRow: curRow, sits:[] };
          acc[blki] = blk;
        }
        if (blk.min > i) blk.min = i;
        if (blk.max < i) blk.max = i;
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
      ...b,
      cols: b.max - b.min + 1,
      rows: b.maxRow - b.minRow + 1,
      sits: b.sits.map(s => ({
        col: s.col - b.min,
        row: s.row - b.minRow,
      }))
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

async function myFunction() {
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
  const eventArys = await ebFetch('https://www.eventbriteapi.com/v3/organizations/544694808143/events/?name_filter=' + encodeURIComponent('ACCCN 北堂中文实体崇拜注册') + '&time_filter=current_future');
  const eventsMapped = eventArys.events.map(e => {
    return {
      id: e.id,
      date: e.start.local.slice(0, 10)
    }
  });
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
  const preSitItem = {
    quantity: 0,
    emails: [],
    names: [],
    pos: 0,
    id: 1
  }
  const names = pages.attendees.reduce((acc, att) => {
    if (att.cancelled) return acc;
    let ord = acc.oid[att.order_id];
    const key = `${att.profile.name}:${att.profile.email}`.toLocaleLowerCase();
    //console.log(`attend ${key} order ${att.order_id} ${att.cancelled}  ${att.status}`);
    const existing = preSits.find(k => k.toLowerCase() === key);
    if (existing) {
      ord = preSitItem;      
    }
    if (!ord) {
      ord = {
        quantity: 0,
        emails: [],
        names: [],
        key,
        pos: acc.ary.length,
        id: acc.ary.length + 1,
      };
      acc.oid[att.order_id] = ord;      
      acc.ary.push(ord);
    }
    ord.quantity++;
    ord.emails.push(att.profile.email);
    ord.names.push(att.profile.name);
    return acc;
  }, {
    ary: [preSitItem], oid: {}
  }).ary;

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




  const blockConfig =
    [
      [8, 10, 10, 10, 10, 10, 10, 10, 10, 10],
      [8, 10, 10, 10, 10, 10, 10, 10, 10, 10],
      [8, 10, 10, 10, 10, 10, 10, 10, 10, 10],
      [8, 10, 10, 10, 10, 10, 10, 10, 10, 10]
    ];
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



  const blockSitsOrig = blockConfig.map((blk, bi) => {
    return blk.map((rowCnt, curRow) => {
      const r = [];
      for (let i = 0; i < rowCnt; i++) {
        r[i] = {
          user: null,
          uiPos: {
            col: blockStarts[bi] + i,
            row: STARTRow + curRow,
          }
        }
      }
      return r;
    });
  });

  const blockSits = pureSitConfig.map((blk, bi) => {
    return blk.sits.map(s => {
      return s.map(r => {
        if (!r) return null;
        return {
          user: null,
          uiPos: {
            col: blockStarts[bi] + r.col,
            row: STARTRow + r.row,
          }
        }
      });
    });
    return blk.sits((rowCnt, curRow) => {
      const r = [];
      for (let i = 0; i < rowCnt; i++) {
        r[i] = {
          user: null,
          uiPos: {
            col: blockStarts[bi] + i,
            row: STARTRow + curRow,
          }
        }
      }
      return r;
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
  const blkMap = ['A','B','C','D']
  const fit = (who, reverse=false) => {
    let fited = false;
    for (let rowInc = 0; rowInc < numRows; rowInc++) {
      const row = reverse ? numRows - rowInc - 1 : rowInc;
      for (let blki = 0; blki < blockSits.length; blki++) {
        if (credentials.ignoreBlocks[blki]) continue;
        const curBlock = blockSits[blki];
        if (!curBlock) continue;
        const curRow = curBlock[row]?.filter(x=>x);
        if (!curRow) break;
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
              side: 'A',              
            }
            fited = true;
            return;
          } else if (side === 'right') {
            let ind = curRow.length - 1;
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
      for (let row = 0; row < numRows; row++) {
        for (let blki = 0; blki < blockSits.length; blki++) {
          const curBlock = blockSits[blki];
          if (!curBlock) continue;
          const curRow = curBlock[row].filter(x=>x);
          let rowTotal = 0;
          for (let i = 0; i < curRow.length; i++) {
            if (!curRow[i].user) rowTotal++;
          }
          if (rowTotal > maxAva) {
            maxAva = rowTotal;
            curMaxRow = curRow;
            curMaxRowNumber = row;
            curMaxRowBlk = blki;
          }
        }
      }      

      if (curMaxRow) {
        let curMax = 0, curStart = -1, curEnd = -1;
        let bestSpacing = null;
        for (let i = 0; i < curMaxRow.length; i++) {
          const curUser = curMaxRow[i].user;
          if (curStart < 0) {
            if (!curUser) {
              curStart = i;
            }
          } else {
            if (curUser) {
              curEnd = i;
              const size = curEnd - curStart;
              if (size > curMax) {
                curMax = size;
                bestSpacing = {
                  start: curStart,
                  end: curEnd - 1,
                  size, 
                }
              }
              curStart = -1;
            }
          }
        }
        if (bestSpacing) {
          if (bestSpacing.size > who.quantity + (siteSpacing * 2)) {
            const left = Math.round((bestSpacing.size - who.quantity) / 2);
            for (let i = 0; i < who.quantity; i++)
              curMaxRow[bestSpacing.start + left + i].user = who;
            who.posInfo = {
              block: blkMap[curMaxRowBlk],
              row: curMaxRowNumber,
              side: 'B',
            }
          }
        }
      }
    }
    return fited;
  };

  names.forEach(n => {
    fit(n, credentials.preSitsBack.filter(b=>n.key === b.toLowerCase()).length);
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
    const client = await gs.getClient('gzprem');
    const sheet = client.getSheetOps(credentials.sheetId);
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
      ['Code', 'Quantity', '','Pos','','', 'Name', 'Email'],
      ...names.map(n => [n.id, n.quantity, n.pos, n.posInfo.block, n.posInfo.side, n.posInfo.row.toString(), n.names.join(','), n.emails.join(',')])
    ];
    
    const userData = userInfo.map(u => {
      return [u[0].toString(), '', '', '', u[1].toString(), '', '', '', '', {type:'userColor', val:u[2]},u[3],u[4],u[5],'', u[6], '', '', '', '', '', '', u[7]];
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

    await sheet.doBatchUpdate(updateData);
  }

}

if (isLocal) {
  myFunction().catch(err => {
    console.log(get(err, 'response.body') || err);    
  });
}