
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const express = require("express");
import { AuthConfig } from "node-sp-auth-config";
import {getItem, addItem} from './models/spTunnel.js';

var webUrl = 'http://db-sp.darbeirut.com/projects/PILOT-S';
var key = 'Test';
var filter = `Title eq '${  key  }'`;
var _item = await getItem(webUrl, 'RLOD', 'Title,Status', '');



const app = express();

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.get('/api/v1/hw',(req, res)=>{
  res.status(200).send(_item);
  }); //'EForms - Hello World from node');

app.post('/api/v1/addItem',(req, res) =>{
    var _cols = { };
    _cols['Title'] = 'Test Title 123';
    _cols['Category'] = 'Shop Drawings';

    var _addItem = addItem(webUrl, 'RLOD', _cols);
    res.status(200).send('Item is added');
  }); 

const port = 8080;
app.listen(port, ()=>{
 console.log(`App is running on port ${port}`);
});