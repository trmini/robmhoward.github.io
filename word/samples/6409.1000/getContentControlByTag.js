Word.run(function (ctx) {
var ccs = ctx.document.contentControls.getByTag("Customer-Address");
ctx.load(ccs);

ctx.sync()
    .then(function () {
        ctx.references.remove(ccs);
        ctx.sync().then(
            function () {
                console.log("Content Control Text: " + ccs.items[0].text);
            }
         );
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
});