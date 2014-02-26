var express    = require('express'),
    app        = express(),
    path       = require('path'),
    multiparty = require('connect-multiparty')(),
    xlsx       = require('node-xlsx'),
    _          = require('underscore'),
    cluster = require("cluster"),
    fs         = require('fs');

var views   = path.join(__dirname, "views");
var types = {
    "xlsx" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "csv" : "text/csv"
};

app.configure(function () {
    app.use(express.cookieParser());
    app.use(app.router);
    app.use(express.static(path.join(__dirname, "public")));
});

app.get('/', function(req, res){
    fs.readFile(path.join(views, "index.html"), function(err, data){
        if (err) throw err;

        res.writeHead(200, {'Content-Type': 'text/html','Content-Length':data.length});
        res.write(data);
        res.end();
    });
});

app.post('/upload', multiparty, function(req, res){
    console.log('-------- upload request:');
    if( req.files.xlsx.type == types["xlsx"] ){
        console.log('xlsx detected, ', 'ready to parse: ', req.files.xlsx.path);
        var obj = xlsx.parse(req.files.xlsx.path);
        var data = [];
        console.log('xlsx parsed: ', obj);
        for ( var i in obj.worksheets[0].data){
            data[i] = [];
            var process = obj.worksheets[0].data[i];
            for ( var j=process.length-1; j>=0; j-- ) {
                if( ! process[j] || typeof process[j] === "undefined" || ! process[j].value || process[j].value == NaN ){ 
                    process[j] = {value:''};
                }
                
                data[i][j*2] = process[j].value;
                data[i][j*2+1] = encodeURIComponent(process[j].value);
            }
        }
        obj = {"worksheets":[{"data":data}]};
        obj.creator = "mujiang.info";
        obj.lastModifiedBy = "mujiang.info";
        obj.created  = new Date();
        obj.modified = new Date();
        console.log(obj);
        res.writeHead(200, {
            'Content-Type': req.files.xlsx.type,
            'Content-Disposition': 'attachment; filename="'+req.files.xlsx.name+'"'
        });
        res.write(xlsx.build(obj));
        res.end();
    } else {
        //错误处理
    }
});

if (cluster.isMaster) {
    console.log("cluster start");
    cluster.fork();

    cluster.on('exit', function(worker, code, signal) {
        console.error("worder "+worker.process.pid+" died");
        cluster.fork();
    });

    cluster.on('listening', function(worker, address) {
        console.log("Server has started to listening "+address.address+':'+address.port);
    });

} else {
    app.listen(8888);
}
