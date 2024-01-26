const express = require('express');
const app =  express();
//const jssRoutes = require('./api/routes/spread');

//app.use('/spread', jssRoutes);

app.use((req, res, next) => {
    res.status(200).json({
      message: 'Hello World!'
    });
 });

module.exports = app;