
const isLocal = true;

const gs = require('./getSheet');
const request = require('superagent');
const creds = require('./credentials.json');
//const sheet = gs.createSheet();
//const cardStatementData = await sheet.readSheet('A', 'ChaseCard!A:F')

async function myFunction() {
  let pages = null;
  if (isLocal) {
    //var response = await request.get("https://www.eventbriteapi.com/v3/events/156798329023/attendees").set('Authorization', creds.eventBriteAuth).send();
    //pages = response.body
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
  } else {
    var response = UrlFetchApp.fetch("https://www.eventbriteapi.com/v3/events/156798329023/attendees", {
      headers:
      {
        authorization: creds.eventBriteAuth
      }
    });
    const pages = JSON.parse(response.getContentText())
  }
  const names = (pages.attendees.map((a,id) => ({
    quantity: a.quantity,
    email: a.profile.email,
    name: a.profile.name,
    id,
  })));

  let colors = [[0, 0, 255], [0, 255, 0], [255, 0, 0], [0, 255, 255], [255, 0, 255], [255, 255, 0]];
  let fontColor = ['#ffff00', '#ff00ff', '#00ffff', '#000000', '#000000', '#000000'];
  while (colors.length < names.length) {
    colors.map(c => c.map(c => parseInt(c / 2))).forEach(c => {
      if (c[0] + c[1] + c[2] < 255 + 128) {
        fontColor.push('#ffffff');
      } else {
        fontColor.push('#000000')
      }
      return colors.push(c);
    });
  }

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
  const blockColMaxes = blockConfig.map(r => r.reduce(fMax, 0));
  
  const numCols = blockColMaxes.reduce((acc, r) => acc + r + blockSpacing, 0);
  const numRows = blockConfig.map(r => r.length).reduce(fMax, 0);
  
  const STARTCol = 3;
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



  const blockSits = blockConfig.map((blk, bi) => {
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
  const fit = (who) => {
    let fited = false;
    for (let row = 0; row < numRows; row++) {
      for (let blki = 0; blki < blockSits.length; blki++) {
        const curBlock = blockSits[blki];
        const curRow = curBlock[row];
        ['left', 'right'].forEach(side => {
          if (fited) return;
          if (side === 'left') {
            if (curRow[0].user) return;
            for (let i = 0; i < who.quantity; i++) curRow[i].user = who;
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
          }
        });
      }
    }
    if (!fited) {
      let maxAva = 0;
      let curMaxRow = null;
      for (let row = 0; row < numRows; row++) {
        for (let blki = 0; blki < blockSits.length; blki++) {
          const curBlock = blockSits[blki];
          const curRow = curBlock[row];
          let rowTotal = 0;
          for (let i = 0; i < curRow.length; i++) {
            if (!curRow[i].user) rowTotal++;
          }
          if (rowTotal > maxAva) {
            maxAva = rowTotal;
            curMaxRow = curRow;
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
          }
        }
      }
    }
    return fited;
  };

  names.forEach(fit);


  if (!isLocal) {
    const sittingSheet = SpreadsheetApp.openById('1p7W0Gwh88tCSiEA7Y6S_tVfyTMMbzotjaSZmfgEpDHY');

    const sheet = sittingSheet.getSheets()[0];
    sheet.clear();
    sheet.setColumnWidths(STARTCol, numCols, CELLSIZE);
    sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
    blockSits.forEach(blk => {
      blk.forEach(row => {
        row.forEach(c => {
          range.setValue(c.user ? c.user.id : '-');
          if (c.user) {
            range.setBackground(colors[c.user.id]);
            range.setFontColor(fontColor[c.user.id]);
          }
        });
      })
    });
    const userInfo = [
      ['Name', 'Quantity', 'Email'],
      ...names.map(n => [n.name, n.quantity, n.email])
    ];
    sheet.getRange(namesStartRow, 1, names.length + 1, 3).setValues(
      userInfo
    )

  } else {
    const sheet = gs.createSheet();    
    const data = [];
    for (let i = 0; i < numRows; i++) data[i] = [];
    blockSits.forEach(blk => {
      blk.forEach(r => {
        r.forEach(c => {
          data[c.uiPos.row - STARTRow][c.uiPos.col - STARTCol] = c.user ? c.user.id: 'e';
        })
      })
    })
    await sheet.updateSheet('1p7W0Gwh88tCSiEA7Y6S_tVfyTMMbzotjaSZmfgEpDHY', `Sheet1!C3:AV12`, data);    
  }

}

if (isLocal) {
  myFunction().catch(err => {
    console.log(err)
  });
}