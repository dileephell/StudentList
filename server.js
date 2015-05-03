// MEAN Stack RESTful API Tutorial - studentlist

var express = require('express');
var app = express();
var mongojs = require('mongojs');
var db = mongojs('studentlist', ['studentlist']);
var bodyParser = require('body-parser');
var nodeExcel = require('excel-export');
var json2xls = require('json2xls');
var fs = require('fs');

app.use(express.static(__dirname + '/public'));
app.use(bodyParser.json());

app.get('/students', function (req, res) {
  console.log('I received a GET request');

  db.studentlist.find(function (err, docs) {
    console.log(docs);
    res.json(docs);
  });
});

app.post('/students', function (req, res) {
  console.log(req.body);
  db.studentlist.insert(req.body, function(err, doc) {
    res.json(doc);
   
  });
});

app.get('/Excel', function(req, res){
    var conf ={};
  // uncomment it for style example  
  // conf.stylesXmlFile = "styles.xml";
    conf.cols = [{
        caption:'string',
        captionStyleIndex: 1,        
        type:'string',
        beforeCellWrite:function(row, cellData){
             return cellData.toUpperCase();
        }
        , width:15
    },{
        caption:'date',
        type:'date',
        beforeCellWrite:function(){
            var originDate = new Date(Date.UTC(1899,11,30));
            return function(row, cellData, eOpt){
              // uncomment it for style example 
              // if (eOpt.rowNum%2){
                // eOpt.styleIndex = 1;
              // }  
              // else{
                // eOpt.styleIndex = 2;
              // }
              if (cellData === null){
                eOpt.cellType = 'string';
                return 'N/A';
              } else
                return (cellData - originDate) / (24 * 60 * 60 * 1000);
            } 
        }()
        , width:20.85
    },{
        caption:'bool',
        type:'bool'
    },{
        caption:'number',
        type:'number',
        width:30
    }];
    db.studentlist.find(function (err, docs) {
    console.log(docs);
    res.json(docs);
  var result = nodeExcel.execute(conf);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats');
  res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
  res.end(result, 'binary');
  
    
})
});

  

app.delete('/students/:id', function (req, res) {
  var id = req.params.id;
  console.log(id);
  db.studentlist.remove({_id: mongojs.ObjectId(id)}, function (err, doc) {
    res.json(doc);
  });
});

app.get('/students/:id', function (req, res) {
  var id = req.params.id;
  console.log('In get by id'+id);
  db.studentlist.findOne({_id: mongojs.ObjectId(id)}, function (err, doc) {
    res.json(doc);
  });
});

app.put('/students/:id', function (req, res) {
  var id = req.params.id;
  console.log(req.body.name);
  db.studentlist.findAndModify({
    query: {_id: mongojs.ObjectId(id)},
    update: {$set: {name: req.body.name, email: req.body.email, number: req.body.number}},
    new: true}, function (err, doc) {
      res.json(doc);
    }
  );
});

app.listen(3000);
console.log("Server running on port 3000");
