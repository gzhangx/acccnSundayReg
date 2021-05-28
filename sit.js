
const isLocal = true;

const gs = require('./getSheet');
const request = require('superagent');
const creds = require('./credentials.json');
//const sheet = gs.createSheet();
//const cardStatementData = await sheet.readSheet('A', 'ChaseCard!A:F')

async function myFunction() {
  let pages = null;
  if (isLocal) {
    var response = await request.get("https://www.eventbriteapi.com/v3/events/156798329023/attendees").set('Authorization', creds.eventBriteAuth).send();
    pages = response.body
  } else {
    var response = UrlFetchApp.fetch("https://www.eventbriteapi.com/v3/events/156798329023/attendees", {
      headers:
      {
        authorization: creds.eventBriteAuth
      }
    });
    const pages = JSON.parse(response.getContentText())
  }
  const names = (pages.attendees.map(a => ({
    quantity: a.quantity,
    email: a.profile.email,
    name: a.profile.name
  })));


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
  console.log('blockColMaxes=')
  console.log(blockColMaxes)
  const numCols = blockColMaxes.reduce((acc, r) => acc + r + blockSpacing, 0);
  const numRows = blockConfig.map(r => r.length).reduce(fMax, 0);
  console.log(`num cols = ${numCols}, rows=${numRows}`);
  const STARTCol = 3;
  const STARTRow = 3;
  const namesSpacking = 3;

  const namesStartRow = STARTRow + numRows + namesSpacking;
  const CELLSIZE = 20;
  const blockStars = blockColMaxes.reduce((acc, b) => {
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

  console.log(`blcick starts`);
  console.log('test1111');

  const blockSits = blockConfig.map((blk, bi) => {
    return blk.map((rowCnt, curRow) => {
      const r = [];
      for (let i = 0; i < rowCnt; i++) {
        r[i] = {
          user: null,
          uiPos: {
            col: blockStars[bi] + i,
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
  blockStars.forEach((bs,i)=>{
    const headers = [];
    for (let j = 0; j < blockColMaxes[i];j++) headers[j] = '';  
    sheet.getRange(STARTRow, blockStars[i], 1, blockColMaxes[i]).setValues([headers]);
    sheet.setColumnWidths(blockStars[i], blockColMaxes[i], CELLSIZE);
    sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
    sheet.getRange(STARTRow,blockStars[i], numRows, blockColMaxes[i]).setBackground('yellow');
  });
  */
  
  
  //sheet.getRange(1,1,1,numCols + STARTCol).setValues([headers]);
  //sheet.setColumnWidths(STARTCol, numCols, CELLSIZE);
  //sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
  //sheet.getRange(STARTRow,STARTCol, numRows, numCols).setBackground('yellow');



  const fit = () => {

  };

  fit();


  if (!isLocal) {
    const sittingSheet = SpreadsheetApp.openById('1p7W0Gwh88tCSiEA7Y6S_tVfyTMMbzotjaSZmfgEpDHY');

    const sheet = sittingSheet.getSheets()[0];
    sheet.clear();
    sheet.setColumnWidths(STARTCol, numCols, CELLSIZE);
    sheet.setRowHeights(STARTRow, numRows, CELLSIZE);
    blockSits.forEach(blk => {
      blk.forEach(row => {
        row.forEach(c => {
          sheet.getRange(c.uiPos.row, c.uiPos.col).setValue('-');
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

  }

}

if (isLocal) {
  myFunction().catch(err => {
    console.log(err)
  });
}