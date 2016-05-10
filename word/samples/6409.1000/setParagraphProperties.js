Word.run(function (ctx) {
var paras = ctx.document.body.paragraphs;
ctx.load(paras);

ctx.sync()
    .then(function () {
        var par = paras.items[0];
        par.lineSpacing = 36;

        ctx.load(par);
        var val = par.lineSpacing;

        ctx.references.remove(paras);
        ctx.sync()
            .then(function () {
                console.log("Success! Setting paragraph line spacing to " + val);
            });
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
})