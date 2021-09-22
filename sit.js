
const isLocal = true;

//const gs = require('./getSheet');
//const fs = require('fs');
//const request = require('superagent');
const { get, sortBy } = require('lodash');
const { toSimp } = require('./gbTran');

const perfixOnLastName = true;
const sheetName = 'Sheet1';
const credentials = require('./credentials.json');
const utils = require('./util');
async function myFunction() {

  /* current saved
主席領詩	C0
司琴	A1
帶位	B11
带位	B11
敬拜	B0
投影	D8
音效	D8
牧師	C0
IT 執事	D9
诗班	B0
執事	D9
領詩	C0
講員	C0
詩班	B0
领诗	B0
  */  


  const initInfo = await utils.initAll();
  const nextSundays = initInfo.nextSundays;
  const nextSunday = nextSundays[0];

//const client = await gs.getClient('gzprem');
  const sheet = initInfo.sheet; //client.getSheetOps(credentials.sheetId);
  const fixedInfoOrig = await sheet.readValues(`'${nextSunday}'!A1:F300`).catch(err => {
    console.log('Unable to load fixed')
    //console.log(err.response.body);
    return [];
  });
  const preFixesInfo = [];
  //(await sheet.readValues(`'PreFixes'!A1:D300`)).map(v => {
  //  return {
  //    prefix: v[0],
  //    pos: v[1],
  //  }
  //});

  const templates = initInfo.templates;

  const debugComplted = !!templates.filter(f => f[0] === 'debugCompleted' && f[1] === 'TRUE').length;
  if (debugComplted) {
    console.log(`Warning, debug mode`);
  }
  const ebQueryStatus = {
    time_filter: debugComplted ? 'past' : 'current_future',
    status: debugComplted ? 'completed' : 'live'
  }



  const PREASSIGNEDSIT_ARYNAME = 'pprefixes';
  const { preAssignedSits } = templates.filter(f => f[0] === 'assignedPrefixes').reduce((acc, f) => {
    const prefixes = toSimp(f[1]).split(',');
    const posRaw = f[2].split('-');
    const blk = posRaw[0][0];
    const nopack = f[3] === 'nopack';
    posRaw[0] = posRaw[0].slice(1);
    if (posRaw.length === 1) posRaw.push(posRaw[0]);
    const from = posRaw[0];
    const to = posRaw[1];
    acc.preAssignedSits[blk] = acc.preAssignedSits[blk] || {};
    prefixes.forEach(prefix => {
      preFixesInfo.push({
        prefix,
        pos: `${blk}${from}`
      })
    });
    if (nopack) return acc;
    for (let i = from; i <= to; i++) {
      const blki = acc.preAssignedSits[blk][i] || {
        [PREASSIGNEDSIT_ARYNAME]:[]
      };
      acc.preAssignedSits[blk][i] = blki;
      prefixes.forEach(pf => {
        blki[PREASSIGNEDSIT_ARYNAME].push(pf);
      })
    }
    return acc;
  }, {
    preAssignedSits: {},
  });

  const emailToFuncMappings = templates.filter(f => f[0] === 'mapping' && f[1] && f[2]).reduce((acc, f) => {
    acc[f[2]] = f[1];
    return acc;
  }, {});

  const parseIgnoreBlocks = () => {
    const row = templates.filter(f => f[0] === 'ignoreBlocks')[0];
    if (!row) return [];
    try {
      return JSON.parse(row[1]);
    } catch (err) {
      console.log(`failed to parse ignoreBlocks ${row[1]}`);      
    }
    return [];
  }
  const ignoreBlocks = parseIgnoreBlocks();


  const { pureSitConfig, getDisplayRow, CELLSIZE, blkLetterToId, numRows, colNumDisplay } = initInfo;
//console.log(pureSitConfig.map(s=>({cols: s.cols, rows: s.rows})))
//return console.log(pureSitConfig.map(r => r.sits.map(v => v.map(vv => vv ? 'X' : ' ').join('')).join('\n')).join('\n'));

  const authorizationToken = credentials.eventBriteAuth;
  const ebFetch = async (url, prms) => {
    if (prms) {
      url = url + '?' + Object.keys(prms).map(n => `${n}=${encodeURIComponent(prms[n])}`).join('&');
    }
    //console.log(`url=${url}`);
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

  const searchTitle = get(templates.filter(t => t[0] === 'searchTitle'),[0,1]) || credentials.eventTitle;
  //console.log(`trying to search event ${searchTitle}`);
  const eventArys = await ebFetch('https://www.eventbriteapi.com/v3/organizations/544694808143/events/',
    { name_filter: searchTitle, time_filter: ebQueryStatus.time_filter }
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

  const eventName = nextGoodEvent.name.text;
  console.log(`Event name ${eventName}`);
  
  let attendees = [];
  let attendeesPrms = null;
  while (true) {
    const pages = await ebFetch(`https://www.eventbriteapi.com/v3/events/${nextGoodEvent.id}/attendees`, attendeesPrms);
    attendees = attendees.concat(pages.attendees.filter(x=>!x.cancelled));
    if (pages.pagination.has_more_items) {
      attendeesPrms = {
        continuation: pages.pagination.continuation,
      }
      continue;
    }
    break;
  }
  console.log(`Total attendees ${attendees.length}`);
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
  

  const assignedByEmail = templates.filter(f => f[0] === 'assignedByEmail').reduce((acc, f) => {
    acc[f[1].toLocaleLowerCase()] = f[2]; //email:A#
    return acc;
  }, {});
  const fixedInfo = fixedInfoOrig.concat(
    attendees.filter(att => assignedByEmail[att.profile.email.toLocaleLowerCase()]).map(att => {
      const sitLong = assignedByEmail[att.profile.email.toLocaleLowerCase()];
      const dashInd = sitLong.indexOf('-');
      if (dashInd < 0) {
        console.log(`Warning, assigned sit for ${att.profile.email} need to have full sit like A0-3, but got ${sitLong}`);
        throw 'err';
      }      
      const profile = att.profile;
      return [
        att.order_id,
        profile.name,
        profile.email,
        sitLong.substr(0, dashInd),
        sitLong,
        { first_name: toSimp(profile.first_name), last_name: toSimp(profile.last_name), email: profile.email, name: profile.name },
      ]
    })
  )
  const preSits = fixedInfo.reduce((acc, f) => {
    if (f[4])
      acc[f[0]] = f;
    return acc;
  }, {});
  const preSiteItems = fixedInfo.filter(v => v[3]).map((v, pos) => {
    const order_id = v[0];
    const name = toSimp(v[1]);
    const email = v[2];
    const blkRowId = v[4];
    const profile = v[5] ? JSON.parse(v[5]) : {};
    const rc = blkRowId.slice(1).split('-');
    const key = `${name}:${email}`.toLocaleLowerCase();
    return {
      quantity: 1,
      emails: [email],
      names: [name],
      profiles: [profile],
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
    if (!acc[r.blkRowId]) {
      acc[r.blkRowId] = [];
    }
    acc[r.blkRowId].push(r);
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

    const { profile } = att;
    if (!ord) {
      ord = {
        quantity: 0,
        emails: [],
        names: [],
        profiles: [],
        keys: [],
        order_id: att.order_id,
        key,
        pos: acc.ary.length,
        id: acc.ary.length + 1,
        name: toSimp(profile.name),
        email: profile.email,
      };
      acc.oid[att.order_id] = ord;      
      acc.ary.push(ord);
    }
    ord.quantity++;
    
    ord.emails.push(profile.email);
    ord.names.push(toSimp(profile.name));
    const pfInf = {
      first_name: toSimp(profile.first_name),
      last_name: toSimp(profile.last_name),
      email: profile.email,
      name: toSimp(profile.name),
    };
    preFixesInfo.find(p => {
      if (perfixOnLastName) {
        if (pfInf.first_name.startsWith(p.prefix)) {
          pfInf.first_name = pfInf.first_name.substr(p.prefix.length).trim();
          if (!pfInf.last_name.startsWith(p.prefix))
            pfInf.last_name = `${p.prefix} ${pfInf.last_name}`;
        }
      }
    });
    ord.profiles.push(pfInf)
    ord.keys.push(key);
    return acc;
  }, {
    ary: preSiteItems, oid: {}
  }).ary,'order_id');

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


  
  const { blkMap } = initInfo;
  const blockSits = initInfo.generateBlockSits(preSiteItemsByBlkRowId);


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
  const fitSection = (who, sectionName) => {
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
      const blkName = blkMap[blki];
      const preAssignedNamesForSit = get(preAssignedSits, [blkName, row, PREASSIGNEDSIT_ARYNAME]);
      if (preAssignedNamesForSit) {
        const matched = preAssignedNamesForSit.find(pfx => who.name.startsWith(pfx));
        if (!matched) continue;
      } else continue;
      for (let i = 0; i < curRow.length; i++) {
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

  let seated = 0;
  const fit = (who, reverse = false) => {
    if (who.posInfo) return true;
    //const ignoreBlocks = credentials.ignoreBlocks;
    let fited = false;
    for (let rowInc = 0; rowInc < numRows; rowInc++) {      
      const row = reverse ? numRows - rowInc - 1 : rowInc;
      for (let blki = 0; blki < blockSits.length; blki++) {
        if (!pureSitConfig[blki].goodRowsToUse[rowInc]) continue;
        if (ignoreBlocks[blki]) continue;
        const blkName = blkMap[blki];
        const preAssignedNamesForSit = get(preAssignedSits, [blkName, row, PREASSIGNEDSIT_ARYNAME]);
        if (preAssignedNamesForSit) {
          const matched = preAssignedNamesForSit.find(pfx => who.name.startsWith(pfx));
          if (!matched) continue;
        }
        const curBlock = blockSits[blki];
        //if (!curBlock) continue;
        const curRow = curBlock[row]?.filter(x=>x);
        if (!curRow) continue;
        ['left', 'right'].forEach(side => {
          if (fited) return;
          if (side === 'left') {
            let tryCol = 0;
            if (curRow[tryCol].user) return;
            while (curRow[tryCol] && curRow[tryCol].sitTag !== 'X') tryCol++;
            if (!curRow[tryCol]) return false;
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
            seated++;
            fited = true;
            return;
          } else if (side === 'right') {
            let ind = curRow.length - 1;
            if (curRow[ind].user) return;
            while (curRow[ind] && curRow[ind].sitTag !== 'X') {
              ind--;
            }
            if (!curRow[ind]) return;
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
            seated++;
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
          if (ignoreBlocks[blki]) continue;
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
              seated++;
            }
            return true;
          }
        }
      }
    }
    return fited;
  };

  //const choreNames = ['詩 ', '詩-'];
  preFixesInfo.filter(p => p.prefix).forEach(prefixInfo => {
    names.filter(n => n.name.startsWith(prefixInfo.prefix) || emailToFuncMappings[n.email] === prefixInfo.prefix).forEach(n => {
      fitSection(n, prefixInfo.pos);
    });
  });

  let unableToSet = 0;
  let totalSeated = 0;
  names.forEach(n => {
    if (!fit(n)) {
      console.log(`Warning, unable to fit ${n.name}`);
      unableToSet++;
    } else {
      totalSeated++;
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
    
    const data = initInfo.getDisplayData(blockSits);
        
    const endColumnIndex = initInfo.STARTCol + initInfo.numCols;
    //console.log(`end col num=${initInfo.numCols} ${initInfo.STARTCol} end=${endColumnIndex}`);
    const sheetInfos = await sheet.sheetInfo();
    const sheetInfo = sheetInfos.find(s => s.title === sheetName);
    if (!sheetInfo) {
      console.log(`sheet ${sheetName} not found`);
    }
    
    const createSheet = async (name, freeInd) => {
      if (!sheetInfos.find(s => s.title === name)) {
        while (true) {
          if (sheetInfos.find(s => s.sheetId === freeInd)) {
            freeInd++;
            continue;
          }
          break;
        }
        console.log(`freeInd for ${name} ${freeInd}`);
        await sheet.createSheet(freeInd, name);
      }
      return freeInd;
    };

    let freeInd = await createSheet(nextSunday, 0);
    const DisplaySheetId = await createSheet(`${nextSunday}Display`, freeInd + 1);
    
    const { sheetId } = sheetInfo;
    const namesFlattened = sortBy(names.filter(f => f.posInfo).reduce((acc, n) => {
      for (let i = 0; i < n.names.length; i++) {
        const pf = n.profiles[i];
        acc.push({
          ...n,
          namesj: `${pf.last_name} ${pf.first_name}`, //n.names[i],
          emailsj: n.emails[i],
          profile: pf,
        })
      }
      return acc;
    },[]),'namesj');
    const userInfo = [
      ['Code', 'Quantity', '','Pos','','', 'Name', 'Email','ActualPos'],
      ...namesFlattened.filter(f => f.posInfo).map(n => [n.id, n.quantity, n.pos, n.posInfo.block, getDisplayRow(n.posInfo.row).toString(), n.posInfo.side,  n.namesj, n.emailsj, `r=${n.posInfo.rowInfo.row} c=${n.posInfo.rowInfo.col}`])
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
      return [u[0].toString(), '', '', '', u[1].toString(), '', '', '', '', { type: 'userColor', val: u[2] }, u[3], u[4], !u[5]?'':colNumDisplay||u[5], '', u[6], '', '', '', '', '', '', u[7], '', '', '', '', '', '', '', '', '', '', u[uoff +8]];
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
        //console.log(`updating column endColumnIndex=${endColumnIndex} sheetInfo.columnCount=${sheetInfo.columnCount} ${endColumnIndex > sheetInfo.columnCount}` );
        //console.log({
        //  sheetId,
        //  dimension: 'COLUMNS',
        //  length: endColumnIndex - sheetInfo.columnCount,
        //})
        await sheet.doBatchUpdate({ requests });
        //console.log('column updated');
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
    console.log('update next')
    await sheet.updateValues(`'${nextSunday}'!A1:F${userInfo.length + 2}`,
      [[eventName, '', '', '', '']].concat(namesFlattened.map(n => {
        return [n.order_id, n.namesj, n.emailsj, `${n.posInfo.block}${getDisplayRow(n.posInfo.row).toString()}${colNumDisplay || n.posInfo.side}`
        , `${n.posInfo.block}${n.posInfo.rowInfo.row}-${n.posInfo.rowInfo.col}`
        //,get(preSits,[n.order_id,5])||''
        , JSON.stringify(n.profile)
      ];
      })));
    
    
    const printRowsPerPage = 32;
    const resd = userInfo.length % printRowsPerPage;
    const totalLines = (resd > 0 ? printRowsPerPage * 2 : printRowsPerPage) + (parseInt(userInfo.length / printRowsPerPage) * printRowsPerPage);
    
    const hankDspRows = namesFlattened.concat(Array(totalLines - namesFlattened.length).fill({ isEmpty: true, texts: ['', '', ''] }));
    hankDspRows[namesFlattened.length] = { texts: [''] };
    hankDspRows[namesFlattened.length + 1] = { texts: ['补注册登记(Walk In Registration)'] };
    hankDspRows[namesFlattened.length + 2] = { texts: ['姓名' , '', '电邮地址'] };
    const sortByRow = sortBy(namesFlattened, n => `${n.posInfo.block}${n.posInfo.rowInfo.row}-${n.posInfo.rowInfo.col}`);
    /*
    await sheet.updateValues(`'${nextSunday}Display'!A1:F${hankDspRows.length + 1}`,
      hankDspRows.map((n, rown) => {
        if (!sortByRow[rown]) {
          return [n.text || ''];
        }
        //n.order_id, n.namesj, n.emailsj, `${n.posInfo.block}${getDisplayRow(n.posInfo.row).toString()}${colNumDisplay || n.posInfo.side}`
        //  , `${sortByRow[rown].posInfo.block}${sortByRow[rown].posInfo.rowInfo.row + 1}`
        //  , sortByRow[rown].namesj
        return [n.namesj, `${n.posInfo.block}${getDisplayRow(n.posInfo.row).toString()}`
          , `${sortByRow[rown].posInfo.block}${sortByRow[rown].posInfo.rowInfo.row+1}`
          , sortByRow[rown].namesj
        ];
      }));
    
    */
    
    const lastBatchUpdateData = createCellRequest({
      hankDspRows,
      mergeRowStartIndex: namesFlattened.length,
      sheetId: DisplaySheetId,
      endColumnIndex: 7,
      endRowIndex: hankDspRows.length + 1,
      rows: hankDspRows.map((n, rown) => {
        const borderStyle1 = {
          style: 'SOLID',
          width: 1,
          color: {
            blue: 0,
            green: 0,
            red: 0
          }
        };
        if (!sortByRow[rown]) {
          const isCenter = n.texts?.length == 1;
          return {
            values: n.texts.map(v=>createCellRowData({
              stringValue: v,
              horizontalAlignment: isCenter ? 'CENTER' : 'Left',
              bold: isCenter,
              borders: {
                left: borderStyle1,
                bottom: borderStyle1,
                right: borderStyle1,
              }
            })),
          };
        }
        return {
          values: [rown.toString(), n.namesj, `${n.posInfo.block}${getDisplayRow(n.posInfo.row).toString()}`, ''
            ,'', `${sortByRow[rown].posInfo.block}${sortByRow[rown].posInfo.rowInfo.row + 1}`
            , sortByRow[rown].namesj
          ].map((stringValue, pos) => {
            
            const borderStyle2 = {
              ...borderStyle1,
              width: 3
            };
            let borders = {
              bottom: borderStyle1,
              left: borderStyle1,
              right: borderStyle1,
            };            
            if (pos == 3) {
              borders.right = borderStyle2;
            }
            return createCellRowData({
              stringValue,
              borders,
            })
          })
        };
      }),
    });
    const lastRes = await sheet.doBatchUpdate(lastBatchUpdateData);


    console.log(lastRes)

    //await utils.sendEmail();
  }

  return {
    seated,
    unableToSet,
    totalSeated,
  }
}


function createCellRequest({
  sheetId,
  endColumnIndex,
  endRowIndex,
  mergeRowStartIndex,
  rows,
  hankDspRows,
}) {
  const singleCells = hankDspRows.map((x, pos) => {
    if (!x.texts || x.texts.length !== 1) return null;
    return {
      mergeCells: {
        range: {
          sheetId,
          startColumnIndex: 0,
          endColumnIndex,
          startRowIndex: pos,
          endRowIndex: pos + 1,
        },
        mergeType: 'MERGE_ROWS',
      }
    }
  }).filter(x => x);
  
  const TWODIGITCELLSIZE = 27;
  const DISPLAYNAMECELLSIZE = 200;
  const updateData = {
    requests: [
      {
        unmergeCells: {
          range: {
            sheetId,
            startColumnIndex: 0,
            endColumnIndex,
            startRowIndex: 0,
            endRowIndex,
          }
        },
      },
      ...singleCells,
      {        
        mergeCells: {
          range: {
            sheetId,
            startColumnIndex: 0,
            endColumnIndex:2,
            startRowIndex: mergeRowStartIndex + 2,
            endRowIndex ,
          },
          mergeType: 'MERGE_ROWS',
        },
      },
      {
        mergeCells: {
          range: {
            sheetId,
            startColumnIndex: 2,
            endColumnIndex,
            startRowIndex: mergeRowStartIndex + 2,
            endRowIndex,
          },
          mergeType: 'MERGE_ROWS',
        },
      },
      {
        //display count column
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 0,
            endIndex: 1
          },
          properties: {
            pixelSize: TWODIGITCELLSIZE
          },
          fields: 'pixelSize'
        }
      },
      {
        //display name first
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 1,
            endIndex: 2
          },
          properties: {
            pixelSize: DISPLAYNAMECELLSIZE
          },
          fields: 'pixelSize'
        }
      },
      {
        //display name first pos
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 2,
            endIndex: 3
          },
          properties: {
            pixelSize: TWODIGITCELLSIZE
          },
          fields: 'pixelSize'
        }
      },
      {
        //display name first check
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 3,
            endIndex: 4
          },
          properties: {
            pixelSize: TWODIGITCELLSIZE
          },
          fields: 'pixelSize'
        }
      },
      {
        //display name second check
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 4,
            endIndex: 5
          },
          properties: {
            pixelSize: TWODIGITCELLSIZE
          },
          fields: 'pixelSize'
        }
      },
      {
        //display name second pos
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 5,
            endIndex: 6
          },
          properties: {
            pixelSize: TWODIGITCELLSIZE
          },
          fields: 'pixelSize'
        }
      },
      {
        //display name second 
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 6,
            endIndex: 7
          },
          properties: {
            pixelSize: DISPLAYNAMECELLSIZE
          },
          fields: 'pixelSize'
        }
      },
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
          rows,
        }
      }
    ]
  };
  return updateData;
}

function createCellRowData({
  stringValue, //userEnteredString value
  backgroundColor = { red: 1, green: 1, blue: 1 },
  horizontalAlignment = 'Left',
  foregroundColor = { red: 0, green: 0, blue: 0 },
  bold = false,
  fontSize = 10,
  strikethrough = false,
  underline = false,
  italic = false,
  borders = {
    bottom: {
      style: 'SOLID',
      width: 1,
      color: {
        blue: 0,
        green: 0,
        red: 0
      }
    },
    left: {
      style: 'SOLID',
      width: 1,
      color: {
        blue: 0,
        green: 0,
        red: 0
      }
    }
  },
}) {
  const cell = {
    userEnteredValue: { stringValue }
  };
  
  cell.userEnteredFormat = {
    backgroundColor,
    horizontalAlignment,
    textFormat: {
      foregroundColor,
      //fontFamily: string,
      fontSize,
      bold,
      italic,
      strikethrough,
      underline,
    },
    borders,
  };
  return cell;
}

if (isLocal) {
  myFunction()
    .then(res => {
      console.log(`new seat = ${res.seated}, totalSeated=${res.totalSeated}`);
      if (res.unableToSet) {
        console.log(`!!!!!!!!!!!!!!!!Unable to seat ${res.unableToSet}`);
      }
    })
    .catch(err => {
    console.log(get(err, 'response.body') || err);
  });
}