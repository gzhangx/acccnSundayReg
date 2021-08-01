
const isLocal = true;


const utils = require('./util');
async function doWork() {
    const initInfo = await utils.initAll();
    await initInfo.sendEmail();
}
doWork();
//utils.sendEmail();