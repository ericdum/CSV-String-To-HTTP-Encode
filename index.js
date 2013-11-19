var express    = require('express'),
    app        = express(),
    path       = require('path'),
    multiparty = require('connect-multiparty')(),
    xlsx       = require('node-xlsx'),
    _          = require('underscore'),
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
        console.log('xlsx parsed');
        obj.worksheets[0].maxCol *= 2;
        for ( var i in obj.worksheets[0].data){
            var process = obj.worksheets[0].data[i];
            for ( var j=process.length-1; j>=0; j-- ) {
                process[j*2] = _.clone(process[j]);
                if( ! process[j*2] ) process[j]={formatCode: "General"};
                process[j*2+1] = _.clone(process[j]);
                process[j*2+1].value = encodeURIComponent(process[j*2+1].value);
            }
            obj.worksheets[0].data[i] = _.clone(process);
            obj.worksheets[0].length = i+1
        }
        console.log('row count: ', obj.worksheets.length);
        res.writeHead(200, {
            'Content-Type': req.files.xlsx.path,
            'Content-Disposition': 'attachment; filename="'+req.files.xlsx.name+'"'
        });
        res.write(xlsx.build(obj));
        res.end();
    } else {
        //错误处理
    }
});

app.listen(8888);
console.log('Listening on port 8888');
