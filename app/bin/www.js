// #!/usr/bin/env node

/**
 * Module dependencies.
 */

const databaseContextPg = require("database-context-pg");
const connectionSetting = require("../dbconnect");

const connectionConfig = connectionSetting.config;
const condb = new databaseContextPg(connectionConfig);

var app = require('../app');
var debug = require('debug')('webapi:server');
var http = require('http');
const { Server } = require("socket.io");
var _config = require('../appSetting.js');
/**
 * Get port from environment and store in Express.
 */

var port = normalizePort(process.env.PORT || _config.port);
app.set('port', port);

/**
 * Create HTTP server.
 */

var server = http.createServer(app);
var server2 = http.createServer(app);

const io = new Server(server2,{
  cors:{
    origin: '*'
  }
});

var currentConnections = [];
io.on('connect', (socket) => {
  /* var currentConnections = []; */
  //console.log('a user connected',socket.id);
  socket.emit('welcome',"WELCOME FROM SERVER")

  
  socket.on('online',(user_id)=>{
    console.log('User Online -->',user_id)

    let temp = {
      socket_id:socket.id,
      user_id:user_id,
    }
    currentConnections.push(temp);
    updateUserOnline(user_id,true)
    console.log('online',currentConnections)
    for(let item of io.sockets.adapter.rooms){
     /*  console.log(`Socket ${socket.id} leaving ${item[0]}`);
    socket.leave(item[0]); */
    //console.log("room",item[0])

    let project_id = (item[0] || '').toString().split(':').length > 1 ? (item[0] || '').toString().split(':')[1] : null ;
    console.log("project_id",project_id)
    if(project_id){
      io.to(item[0]).emit('online_update', {project_id:project_id });
    }

    /* socket.leave(item); */
  }

  })
  


  socket.on('disconnect', () =>{
      console.log(`Disconnected: ${socket.id}`)
      console.log(currentConnections);
      let indexCon = currentConnections.findIndex((dis)=> dis.socket_id == socket.id)
      if(indexCon > -1){
        updateUserOnline(currentConnections[indexCon].user_id,false)
        currentConnections.splice(indexCon,1);
        for(let item of io.sockets.adapter.rooms){
          /*  console.log(`Socket ${socket.id} leaving ${item[0]}`);
         socket.leave(item[0]); */
         //console.log("room",item[0])
         let project_id = (item[0] || '').toString().split(':').length > 1 ? (item[0] || '').toString().split(':')[1] : null ;
         console.log("project_id",project_id)
         if(project_id){
           io.to(item[0]).emit('online_update', {project_id:project_id });
         }
      /*    io.to(item[0]).emit('online_update', {project_id:project_id }); */
         /* socket.leave(item); */
       }
      }
      console.log(currentConnections);

  }


  
  );
  socket.on('join', (room) => {
  
  console.log(`Socket ${socket.id} joining ${room}`);
  socket.join(room);


// console.log('connect',io.sockets.adapter.rooms)
});
socket.on('leave', (room) => {
/*   console.log(`Socket ${socket.id} leaving ${room}`); */
  for(let item of io.sockets.adapter.rooms){
      //console.log(`Socket ${socket.id} leaving ${item[0]}`);
    socket.leave(item[0]);
    /* socket.leave(item); */
  }
  /* socket.leave(room); */
});
socket.on('chat', (data) => {
      const { message, room } = data;
      //console.log(`msg: ${message}, room: ${room}`);
      //console.log('chat',io.sockets.adapter.rooms)
      io.to(room).emit('chat', data);
   });

/*   socket.on("chat message", (message) => {
    socket.to(roomID).emit("chat message", message);
  }); */

/*   socket.on("disconnect", () => {
    // leave room
    room.leaveRoom(roomID);
  }); */
});

/*   socket.on('chat message', (data) => {
      console.log(data)
      socket.emit('chat message',data); 
    

  }); */


/* var io = Server.listen(80);

io.on('connection', function (socket) {
  console.log('client connected', socket.id);
	setInterval(function () {
		socket.emit('hi', 'hello from server');
	}, 2000);
}); */
/**
 * Listen on provided port, on all network interfaces.
 */
/* server2.listen(7776); */
server.listen(port);
server.on('error', onError);
server.on('listening', onListening);

/**
 * Normalize a port into a number, string, or false.
 */

function normalizePort(val) {
  var port = parseInt(val, 10);

  if (isNaN(port)) {
    // named pipe
    return val;
  }

  if (port >= 0) {
    // port number
    return port;
  }

  return false;
}

/**
 * Event listener for HTTP server "error" event.
 */

function onError(error) {
  if (error.syscall !== 'listen') {
    throw error;
  }

  var bind = typeof port === 'string'
    ? 'Pipe ' + port
    : 'Port ' + port;

  // handle specific listen errors with friendly messages
  switch (error.code) {
    case 'EACCES':
      console.error(bind + ' requires elevated privileges');
      process.exit(1);
      break;
    case 'EADDRINUSE':
      console.error(bind + ' is already in use');
      process.exit(1);
      break;
    default:
      throw error;
  }
}

/**
 * Event listener for HTTP server "listening" event.
 */

function onListening() {
  var addr = server.address();
  var bind = typeof addr === 'string'
    ? 'pipe ' + addr
    : 'port ' + addr.port;
  debug('Listening on ' + bind);
}




 async function updateUserOnline(user_id,is_online){
  try{
    console.log("เข้าา",user_id,is_online)
    await condb.clientQuery(
      `UPDATE users.users
      SET usr_is_online=$2
      WHERE usr_id=$1;`, [
        user_id,
        is_online
  ]);
  return true;
  }catch(error){
    console.log(error)
    return false;
  }
 

 }