var ctx = new Word.RequestContext();
ctx.document.body.clear();

ctx.sync()
    .then(function () {
        console.log("Success");
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
