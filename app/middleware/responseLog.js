const dbConnect = require('../appSetting.js');
const { Client } = require('pg')
const moment = require("moment");
const { v4: uuidv4 } = require('uuid');
const configLog = dbConnect.dbConnect.dbLogConnect;
class ResponseLog {
    async log(model) {
       
        return new Promise(((resolve, reject) => {
            (async () => {
           /*      var client = new Client(configLog) */
               /*  await client.connect(); */
                try {
                    /* await client.query(`INSERT INTO log( id, date, activity) VALUES ($1, $2, $3)`, [model.id, model.date, model.activity]); */
                    resolve(true);
                } catch (e) {
                    reject(e);
                } finally {
                    /* await client.end(); */
                }
            })().catch(e => {
                console.log(e);
                reject(e);
            })
        }))
    }
    get model() {
        return {
            id: uuidv4(),
            date: moment(new Date()),
            activity: {
                status: true,
                path: "",
                parameter: {},
                response: {},
                error: {}
            }
        }
    }
}
module.exports = ResponseLog