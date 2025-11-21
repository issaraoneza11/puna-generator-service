const { Client } = require('pg');
const connectionSetting = require("../dbconnect");
const connectionConfig = connectionSetting.config;

const clientQuery = async (queryConfig, value) => {
    return new Promise(async (resolve, reject) => {
        try {
            (async () => {
                const client = new Client(connectionConfig);
                await client.connect();
                try {
                    const callback = await client.query(queryConfig, value);
                    resolve(callback);
                } catch (e) {
                    reject(e);
                } finally {
                    await client.end();
                }
            })().catch((e) => {
                console.log(e);
                // throw Error(e);
                reject(e);
            });
        } catch (e) {
            reject(e);
        }
    });
};


module.exports = clientQuery;
