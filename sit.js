
const isLocal = true;

const gs = require('./getSheet');
const fs = require('fs');
const request = require('superagent');
const { get, sortBy } = require('lodash');

const debugComplted = true;
const ebQueryStatus = {
  time_filter: debugComplted?'past':'current_future',
  status: debugComplted?'completed':'live'
}

const sheetName = 'Sheet1';
const credentials = require('./credentials.json');
const utils = require('./util');
async function myFunction() {

  /* current saved
主席領詩	D0
司琴	A1
帶位	B11
敬拜	B0
投影	D8
音效	D8
牧師	D0	4
IT 執事	D9
  */  



  const nextSundays = utils.getNextSundays();
  const nextSunday = nextSundays[0];
  console.log(`nextSunday=${nextSunday}`);

const client = await gs.getClient('gzprem');
const sheet = client.getSheetOps(credentials.sheetId);
  const fixedInfo = await sheet.readValues(`'${nextSunday}'!A1:E300`).catch(err => {
    console.log('Unable to load fixed')
    console.log(err.response.body);
    return [];
  });
  const preFixesInfo = (await sheet.readValues(`'PreFixes'!A1:D300`)).map(v => {
    return {
      prefix: v[0],
      pos: v[1],
      colStart: parseInt(v[2] || 0),
      forceFillEnd: v[3],  //if found, block the rest of the sits to that space
    }
  });


const preSits = fixedInfo.reduce((acc,f) => {
  if (f[4])
    acc[f[0]] = f;
  return acc;
}, {});


  const { pureSitConfig, getDisplayRow, CELLSIZE, blkLetterToId, numRows } = utils;
//console.log(pureSitConfig.map(s=>({cols: s.cols, rows: s.rows})))
//return console.log(pureSitConfig.map(r => r.sits.map(v => v.map(vv => vv ? 'X' : ' ').join('')).join('\n')).join('\n'));

  const authorizationToken = credentials.eventBriteAuth;
  const ebFetch = async (url, prms) => {
    if (prms) {
      url = url + '?' + Object.keys(prms).map(n => `${n}=${encodeURIComponent(prms[n])}`).join('&');
    }
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
  const eventArys = await ebFetch('https://www.eventbriteapi.com/v3/organizations/544694808143/events/',
    { name_filter: credentials.eventTitle, time_filter: ebQueryStatus.time_filter }
  );
  const eventsMappedNonFiltered = eventArys.events.map(e => {
    return {
      id: e.id,
      date: e.start.local.slice(0, 10),
      name: e.name,
      status: e.status,
    }
  }).filter(s => s.status === ebQueryStatus.status);
  const eventsMapped = eventsMappedNonFiltered.filter(x => x.date === nextSunday);
  let nextGoodEvent = (eventsMapped.filter(x => x.date === nextSunday))[0];
  //nextGoodEvent = (eventsMappedNonFiltered)[0]; //TODO: fix
  if (!nextGoodEvent) {
    console.log('Next not found');
    console.log(eventsMappedNonFiltered)
    return;
  }
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
  
  let attendees = [];
  let attendeesPrms = null;
  while (true) {
    const pages = await ebFetch(`https://www.eventbriteapi.com/v3/events/${nextGoodEvent.id}/attendees`, attendeesPrms);
    attendees = attendees.concat(pages.attendees);
    if (pages.pagination.has_more_items) {
      attendeesPrms = {
        continuation: pages.pagination.continuation,
      }
      continue;
    }
    break;
  }
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
  
  const preSiteItems = fixedInfo.filter(v => v[3]).map((v, pos) => {
    const order_id = v[0];
    const name = v[1];
    const email = v[2];
    const blkRowId = v[4];
    const rc = blkRowId.slice(1).split('-');
    const key = `${name}:${email}`.toLocaleLowerCase();
    return {
      quantity: 1,
      emails: [email],
      names: [name],
      name,
      order_id,
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
  const names = sortBy(attendees.reduce((acc, att) => {
    if (att.cancelled) return acc;
    let ord = acc.oid[att.order_id];
    const key = `${att.profile.name}:${att.profile.email}`.toLocaleLowerCase();
    //console.log(`attend ${key} order ${att.order_id} ${att.cancelled}  ${att.status}`);
    const existing = preSits[att.order_id];
    if (existing) {
      return acc;
    }
    if (!ord) {
      ord = {
        quantity: 0,
        emails: [],
        names: [],
        keys: [],
        order_id: att.order_id,
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
    ord.keys.push(key);
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


  
  const { blkMap } = utils;
  const blockSits = utils.generateBlockSits(preSiteItemsByBlkRowId);


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



  const siteSpacing = 2;
  const fitSection = (who, sectionName, colStart) => {
    if (who.posInfo) return true;    
    const blki = blkLetterToId[sectionName[0]]; //block B only , //B11
    const getRowFromSection = () => {
      const pt = sectionName.slice(1);
      if (!pt) return 0;
      return parseInt(pt);
    }
    const curBlock = blockSits[blki];
    for (let row = getRowFromSection(); row < numRows; row++) {
      const curRow = curBlock[row]?.filter(x => x);
      if (!curRow) break;
      for (let i = colStart || 0; i < curRow.length; i++) {
        const cri = curRow[i];
        if (!cri) continue;
        if (cri.sitTag !== 'X') continue;
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
            let tryCol = 0;
            if (curRow[tryCol].user) return;
            while (curRow[tryCol].sitTag !== 'X') tryCol++;            
            for (let i = 0; i < who.quantity; i++){
              if (!curRow[tryCol+i].user)
                curRow[tryCol+i].user = who;
              //else return fit(who, reverse); //this needs testing, i.e. we grow out of current row.
              else return false;
            } 
            who.posInfo = {
              block: blkMap[blki],
              row,
              rowInfo: curRow[tryCol],
              side: 'A',              
            }
            fited = true;
            return;
          } else if (side === 'right') {
            let ind = curRow.length - 1;
            if (curRow[ind].user) return;
            while (curRow[ind].sitTag !== 'X') {
              ind--;
            }
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
          if (bestSpacing.size >= who.quantity + (siteSpacing * 2)) {
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

  //const choreNames = ['詩 ', '詩-'];
  preFixesInfo.forEach(prefixInfo => {
    names.filter(n => n.name.startsWith(prefixInfo.prefix)).forEach(n => {
      fitSection(n, prefixInfo.pos, prefixInfo.colStart);
    });
    if (prefixInfo.forceFillEnd) {
      const sectionName = prefixInfo.pos;
      const blki = blkLetterToId[sectionName[0]]; //block B only , //B11
      const getRowFromSection = () => {
        const pt = sectionName.slice(1);
        if (!pt) return 0;
        return parseInt(pt);
      }
      const curBlock = blockSits[blki];
      for (let row = getRowFromSection(); row < prefixInfo.forceFillEnd; row++) {
        const curRow = curBlock[row]?.filter(x => x);
        if (!curRow) break;
        for (let i = prefixInfo.colStart || 0; i < curRow.length; i++) {
          const cri = curRow[i];
          if (!cri) continue;
          if (cri.user) continue;
          cri.user = 'DBGFILL';
        }
      }
    }
  });
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
      ['Order','Code', 'Name', 'Quantity', 'Email'],
      ...names.map(n => [n.id, n.names.join(','), n.quantity, n.emails.join(',')])
    ];
    sheet.getRange(namesStartRow, 1, names.length + 1, 4).setValues(
      userInfo
    )

  } else {
    
    const data = utils.getDisplayData(blockSits);
        
    const endColumnIndex = utils.STARTCol + utils.numCols;
    console.log(`end col num=${utils.numCols} ${utils.STARTCol} end=${endColumnIndex}`);
    const sheetInfos = await sheet.sheetInfo();
    const sheetInfo = sheetInfos.find(s => s.title === sheetName);
    if (!sheetInfo) {
      console.log(`sheet ${sheetName} not found`);
    }
    
    if (!sheetInfos.find(s => s.title === nextSunday)) {
      let freeInd = 0;
      while (true) {
        if (sheetInfos.find(s => s.sheetId === freeInd)) {
          freeInd++;
          continue;
        }
        break;
      }
      console.log(`freeInd ${freeInd}`);
      await sheet.createSheet(freeInd, nextSunday);
    }
    const { sheetId } = sheetInfo;
    const userInfo = [
      ['Code', 'Quantity', '','Pos','','', 'Name', 'Email','ActualPos'],
      ...names.filter(f => f.posInfo).map(n => [n.id, n.quantity, n.pos, n.posInfo.block, getDisplayRow(n.posInfo.row).toString(), n.posInfo.side,  n.names.join(','), n.emails.join(','), `r=${n.posInfo.rowInfo.row} c=${n.posInfo.rowInfo.col}`])
    ];
 
    // console.log('names==>')
    // console.log(names.map(n => {
    //   return {
    //     rowInfof: n.posInfo.rowInfo,
    //     ...n,
    //   }
    // }));
    const uoff = 1;
    const userData = userInfo.map(u => {
      return [u[0].toString(), '', '', '', u[1].toString(), '', '', '', '', { type: 'userColor', val: u[2] }, u[3], u[4], u[5], '', u[6], '', '', '', '', '', '', u[7], '', '', '', '', '', '', '', '', '', '', u[uoff +8]];
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
    await sheet.updateValues(`'${nextSunday}'!A1:E${userInfo.length + 1}`, names.filter(n=>n.posInfo).map(n => {      
      return [n.order_id,n.names.join(','), n.emails.join(','), `${n.posInfo.block}${getDisplayRow(n.posInfo.row).toString()}${n.posInfo.side}`
        , `${n.posInfo.block}${n.posInfo.rowInfo.row}-${n.posInfo.rowInfo.col}`
      ];
    }));

    //await utils.sendEmail();
  }

}

if (isLocal) {
  myFunction().catch(err => {
    console.log(get(err, 'response.body') || err);    
  });
}