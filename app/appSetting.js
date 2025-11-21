const config = require("./appSettingSite.js");
var _config ={
    development:{
    
    "dbConnect":{
        "user":config.DB_USERNAME_DEV,
        "host": config.DB_SERVER_DEV,
        "database": config.DB_NAME_DEV,
        "password": config.DB_PASSWORD_DEV,
        "port":config.DB_PORT_DEV,
        "max": 10,
        "idleTimeoutMillis": 30000,
        "connectionTimeoutMillis": 5000,
    },
    "dbLogConnect":{
        "user":config.LOG_DB_USERNAME_DEV,
        "host": config.LOG_DB_SERVER_DEV,
        "database": config.LOG_DB_NAME_DEV,
        "password": config.LOG_DB_PASSWORD_DEV,
        "port":config.LOG_DB_PORT_DEV,
        "max": 10,
        "idleTimeoutMillis": 30000,
        "connectionTimeoutMillis": 5000
    }
} ,
production:{
    
    "dbConnect":{
        "user":config.DB_USERNAME_PROD,
        "host": config.DB_SERVER_PROD,
        "database": config.DB_NAME_PROD,
        "password": config.DB_PASSWORD_PROD,
        "port":config.DB_PORT_PROD,
        "max": 10,
        "idleTimeoutMillis": 30000,
        "connectionTimeoutMillis": 5000,
    },
    "dbLogConnect":{
        "user":config.LOG_DB_USERNAME_PROD,
        "host": config.LOG_DB_SERVER_PROD,
        "database": config.LOG_DB_NAME_PROD,
        "password": config.LOG_DB_PASSWORD_PROD,
        "port":config.LOG_DB_PORT_PROD,
        "max": 10,
        "idleTimeoutMillis": 30000,
        "connectionTimeoutMillis": 5000
    }
} ,
 

}

module.exports ={
    "dbConnect":_config[config.START_PROJECT],
    "FTPConnect": {
        "host": config.FTP_CONNECT_HOST ,
        "user": config.FTP_CONNECT_USER ,
        "password": config.FTP_CONNECT_PASSWORD ,
        "remotePath": config.FTP_CONNECT_REMOTE_PATH,
        "localPath": config.FTP_CONNECT_LOCAL_PATH
   },
   "host":config.HOST,
   "port":config.PORT,
   "jwtSecret":config.JWTSECRET,
   "customHeaderKey":config.CUSTOMHERDERKEY,
   "userSwagger":config.USERSWAGGER,
   "passwordSwagger":config.PASSWORDSWAGGER,
   "fixData":{
       "material_unit":{
           "Piece":config.FIXDATA_MATERIAL_UNIT_PIECE,
           "Weight":config.FIXDATA_MATERIAL_UNIT_WEIGHT
       }
   },
   "logAccessPath":config.LOGACCESSPATH
   ,
   "passwordAcademy":config.PASSWORDACADEMY


    }