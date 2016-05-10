Word.run(function (ctx) {

var mySections = ctx.document.sections;
ctx.load(mySections);
ctx.references.add(mySections);

ctx.sync()
    .then(function () {
        var myHeader = mySections.items[0].getHeader("primary");
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        myHeader.insertContentControl();

        ctx.sync()
        .then(function () {
            ctx.references.remove(mySections);
            console.log("Success");
        });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
});