
const {google}=require('googleapis');
const readline=require('readline');


/**
 * 
 * @param {*} ggcredentials {client_id,project_id,auth_uri,client_secret}
 * @param {*} getSaveToken function to save token
 */
function createSheetBase(ggcredentials, userName, getToken, saveToken) {

  // If modifying these scopes, delete token.json. //'https://www.googleapis.com/auth/spreadsheets.readonly'
  const SCOPES=['https://www.googleapis.com/auth/spreadsheets'];

  /**
   * Create an OAuth2 client with the given credentials, and then execute the
   * given callback function.
   * @param {Object} ggcredentials The authorization client credentials.
   * @param {function} callback The callback to call with the authorized client.
   */
  async function authorize(spreadsheetId) {
    //const { client_secret, client_id, redirect_uris } = ggcredentials; //credentials.googleSheet.installed;
    const oAuth2Client=getGoogleClient();

    let token=getToken(userName);
    // Check if we have previously stored a token.  
    if(!token) token=await getNewToken(oAuth2Client);

    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  function getGoogleClient() {
    const {client_secret,client_id}=ggcredentials; //credentials.googleSheet.installed;
    const oAuth2Client=new google.auth.OAuth2(
      client_id,client_secret,'urn:ietf:wg:oauth:2.0:oob');
    return oAuth2Client;
  }
  function authorize1GetUrl() {
    const oAuth2Client=getGoogleClient();
    // Check if we have previously stored a token.
    const authUrl=oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: SCOPES,
    });
    return authUrl;
  }

  async function authorize2CodeToToken(spreadsheetId,code) {
    const oAuth2Client=getGoogleClient();
    return new Promise((resolve,reject) => {
      oAuth2Client.getToken(code,(err,token) => {
        if(err) {
          console.error('Error while trying to retrieve access token',err);
          return reject(err);
        }
        if(saveToken) {
          saveToken(userName,token);
        }
        return resolve(token);
      });
    });
  }


  /**
   * Get and store new token after prompting for user authorization, and then
   * execute the given callback with the authorized OAuth2 client.
   * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
   * @param {getEventsCallback} callback The callback for the authorized client.
   */
  async function getNewToken() {
    return new Promise((resolve,reject) => {
      const authUrl=authorize1GetUrl();
      console.log(`Authorize this app by visiting this url: ${authUrl}`);
      const rl=readline.createInterface({
        input: process.stdin,
        output: process.stdout,
      });
      rl.question('Enter the code from that page here: ',code => {
        rl.close();
        authorize2CodeToToken(spreadsheetId,code).then(token => {
          resolve(token);
        }).catch(err => {
          reject(err);
        })
      });
    });
  }

  async function getSheet(spreadsheetId) {
    //const content = credentials;
    // Authorize a client with credentials, then call the Google Sheets API.
    const auth=await authorize(spreadsheetId);
    const sheets=google.sheets({version: 'v4',auth});
    return sheets;
  }





  async function readSheet(spreadsheetId,range) {
    // Load client secrets from a local file.  
    const sheets=await getSheet(spreadsheetId);
    return new Promise((resolve,reject) => {
      sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      },(err,res) => {
        if(err) {
          console.log('The API returned an error: '+err);
          return reject(err);
        }
        return resolve(res);
      });
    });
  }


  async function readRanges(spreadsheetId,ranges) {
    // Load client secrets from a local file.  
    const sheets=await getSheet(spreadsheetId);
    return new Promise((resolve,reject) => {
      sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges,
      },(err,res) => {
        if(err) {
          console.log('The API returned an error: '+err);
          return reject(err);
        }
        return resolve(res);
      });
    });
  }

  async function updateSheet(spreadsheetId,range,values,valueInputOption='RAW') {
    const sheets=await getSheet(spreadsheetId);
    return sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption,
      resource: {
        values
      }
    }).then(response => {
      var result=response.data;
      //console.log(`${result} cells updated.`);
      return result;
    });
  }


  async function appendSheet(spreadsheetId,range,values,valueInputOption='RAW') {
    const sheets=await getSheet(spreadsheetId);
    return sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption,
      resource: {
        values
      }
    }).then(response => {
      var result=response.data;
      //console.log(`${result} cells updated.`);
      return result;
    });
  }

  return {
    authorize1GetUrl,
    authorize2CodeToToken,

    getSheet,
    readSheet,
    readRanges,
    updateSheet,
    appendSheet,
  }

}

module.exports={
  createSheetBase,  
};