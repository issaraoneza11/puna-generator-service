
const ResponseLog = require("./responseLog");

class Responsedata {
    constructor(req, res) {
        this.req = req;
        this.res = res;
        this.init()
    }

    async init() {
        this.logService = new ResponseLog();
        this.log = this.logService.model;
        const authHeader = this.req.headers.authorization;
        const token = authHeader && authHeader.split(' ')[1]
        this.log.activity.parameter = {
            "body": this.req.body,
            "query": this.req.query,
            "params": this.req.params,
            "header": this.req.headers,
            "payload": token ? authRouter.getPayload(token) : null
        };
        this.log.activity.path = this.req.baseUrl + this.req.path;
    }
    getPayloadData() {
        return this.log.activity.parameter.payload
    }
    async success(data, status = 200) {
        this.log.activity.response = data;
        this.logService.log(this.log).then(res => console.log("save log")).catch(e => console.log(e.message))
        return this.res.status(status).send(data)
    }

    async error(error, response = 400) {
        this.log.activity.error = error
        this.log.activity.status = false;
        this.logService.log(this.log).then(res => console.log("save log")).catch(e => console.log(e.message))
        return this.res.status(400).send({
            response,
            error
        })
    }
}

module.exports = Responsedata;