const querystring = require("querystring");
//rootUrl = 'https://accounts.google.com/o/oauth2/v2/auth';
//opts = { access_type: 'offline', scope: 'https://www.googleapis.com/auth/spreadsheets', response_type: 'code', client_id: 'client_id', redirect_uri: 'urn:ietf:wg:oauth:2.0:oob' }
//return rootUrl + '?' + querystring.stringify(opts);
//'https://accounts.google.com/o/oauth2/v2/auth?access_type=offline&scope=https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fspreadsheets&response_type=code&client_id=client_id&redirect_uri=urn%3Aietf%3Awg%3Aoauth%3A2.0%3Aoob'
//https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=id enable
const request = require('superagent');
const get = require('lodash/get');
const credsJson = require('./credentials.json')

async function doRefresh(creds) {
  if (!creds) return null;
  const { refresh_token, client_id, client_secret } = creds;
  const refreshBody = await request.post('https://oauth2.googleapis.com/token').type('form').send({
    client_secret,
    client_id,
    refresh_token,
    grant_type: 'refresh_token'
  });
  const {
    access_token, expires_in, token_type
  } = refreshBody.body;
  const doOp = (op, id, postFix, data) => request[op](`https://sheets.googleapis.com/v4/spreadsheets/${id}${postFix}`)
    .auth(access_token, { type: 'bearer' })
    .send(data).then(r => r.body);
  const doPost = (id, postFix, data) => doOp('post', id, postFix, data);
  const doBatchUpdate = async (id, data) => doPost(id, ':batchUpdate', data);
  const append = async ({ id, range }, data, opts) => {
    if (!opts) {
      opts = {}
    }
    if (!opts.valueInputOption) opts.valueInputOption = 'USER_ENTERED';
    return await doPost(id, `/values/${range}:append?${querystring.stringify(opts)}`, { values: data });
  };
  const read = async ({ id, range }) => doOp('get', id, `/values/${range}`);
  const getInfo = async id => doOp('get', id, '');
  return {
    access_token,
    expires_on: new Date().getTime() + (expires_in * 1000 - 2000),
    token_type,
    doBatchUpdate,
    append,
    read,
    getSheetOps: id => {
      return {
        doBatchUpdate: data => doBatchUpdate(id, data),
        append: (range, data, ops) => append({ id, range }, data, ops),
        read: range => read({ id, range }),
        readValues: range => read({ id, range }).then(r => get(r, 'values', [])),
        info: () => getInfo(id),
        sheetInfo: async () => {
          const sheetInfos = await getInfo(id); 
          return sheetInfos.sheets.map(s => {
            const props = s.properties;
            return {
              sheetId: props.sheetId,
              title: props.title,
              index: props.index, //not important,
              ...props.gridProperties, //rowCount, columnCount
            }
          })          
        },
        createSheet: async (sheetId, title) => {
          return doBatchUpdate(id, {
            requests: [
              {
                addSheet: {
                  properties: {
                    sheetId,
                    title,
                  }
                }
              }
            ]
          })
        },
        updateValues: (range, values, opts) => {
          if (!opts) {
            opts = {}
          }
          if (!opts.valueInputOption) opts.valueInputOption = 'USER_ENTERED';
          return doOp('put', id, `/values/${range}?${querystring.stringify(opts)}`, {
            values,
          })
        }
      }
    }
  }
}

const clients = {};
async function getClientFor(name, clientCredGetter) {
  let client = clients[name];
  const now = new Date().getTime();
  if (!client || client.expires_on <= now) {
    client = doRefresh(clientCredGetter(name));
    if (!client) return null;
    clients[name] = client;
  }
  return client;
}

async function getClient(name) {
  return getClientFor(name, name => credsJson.googleSheet[name]);
}

async function testinit(name = 'bibleSender') {
  const ccc = credsJson.googleSheet[name];
  console.log(ccc);
  const { client_secret, client_id } = ccc;
  const rrr = await request.post(`https://oauth2.googleapis.com/token`).type('form').send({
    code: '4/1AX4XfWgNG-qBGcQ1yyJZYNJs_5XyU0XRkipQLb89t6jfy0DWCLplxt_b-8Q',
    redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
    client_id,
    client_secret,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    grant_type: 'authorization_code',
  }).then(r => r.body);
  console.log(rrr);
  return;
}
//return testinit();
async function test() {
  const cli = await getClient('gzprem');

  const id = '1MO27odjCsxk6MWL0DygubU53hrtt3OB8SEnqjpUHJ-U';
  await cli.doBatchUpdate(id, {
    "requests": [
      {
        "updateDimensionProperties": {
          "range": {
            "sheetId": 0,
            "dimension": "COLUMNS",
            "startIndex": 0,
            "endIndex": 1
          },
          "properties": {
            "pixelSize": 160
          },
          "fields": "pixelSize"
        },

      }
    ]
  })
  const upres = await cli.doBatchUpdate(id, {
    "requests": [
      {

        "updateCells": {
          "fields": "*",
          "range": {
            "sheetId": 0,
            "startColumnIndex": 0,
            "endColumnIndex": 10,
            "startRowIndex": 0,
            "endRowIndex": 10
          },
          "rows": [
            {
              "values": [
                {
                  "userEnteredFormat": {
                    "backgroundColor": {
                      "blue": 10,
                      "green": 10,
                      "red": 255
                    },
                    "borders": {
                      "bottom": {
                        "style": "SOLID",
                        "width": 8,
                        "color": {
                          "blue": 0,
                          "green": 255,
                          "red": 0
                        }
                      }
                    }
                  },
                  "userEnteredValue": { "stringValue": "strstsdfasdf" }
                },
                {
                  "userEnteredValue": { "stringValue": "col1" }
                }
              ]
            }
          ]
        }
      }
    ]
  });
  console.log(upres);
  const appres = await cli.append({
    id,
    range: `'Sheet1'!A1:B2`
  }, [
    ['aaa', 'bbb1']
  ])
  console.log('append');
  console.log(appres);

  const rres = await cli.read({
    id,
    range: 'A1:B4'
  });
  console.log(rres);



  const sheet = cli.getSheeOps(id);
  sheet.doBatchUpdate({
    "requests": [
      {
        "updateDimensionProperties": {
          "range": {
            "sheetId": 0,
            "dimension": "COLUMNS",
            "startIndex": 0,
            "endIndex": 1
          },
          "properties": {
            "pixelSize": 100
          },
          "fields": "pixelSize"
        },

      }
    ]
  })
  await sheet.append('A:B', [['c', 'D']]);
  console.log(await sheet.read('A1:B4'));
}

//test().catch(err => {
//   console.log(err.response.text);
//})

module.exports = {
  getClient,
}