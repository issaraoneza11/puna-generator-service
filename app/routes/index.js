var express = require('express');
var router = express.Router();
/* var baseService = require('../service/baseService'); */
const _config = require('../appSetting.js');
/* const {
    route
} = require('./users'); */
const path = require('path');
const fs = require('fs');
/* var _baseService = new baseService(); */

/* GET home page. */
router.get('/', function (req, res, next) {
    res.render('index', {
        title: 'Report Service'
    });
});

router.get('/TestConnect2', function (req, res, next) {

    _baseService.TestConnnect2().then(_res => {
        res.status(200).json(_res.rows)
    }).catch(_error => {
        res.status(400).send({
            message: _error.message
        })
    })

});

// a middleware function with no mount path. This code is executed for every request to the router



module.exports = router;