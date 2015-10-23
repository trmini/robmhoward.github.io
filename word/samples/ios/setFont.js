Word.run(function (ctx) {
var paras = ctx.document.body.paragraphs;
ctx.load(paras, { expand: "font" });

ctx.sync()
    .then(function () {
        var font = paras.items[0].font;
        font.size = 32;
        font.bold = true;
        font.color = "#0000ff";
        font.highlightColor = "#ffff00";

        ctx.references.remove(paras);
        ctx.sync()
            .then(function () {
                console.log("Success");
            }
        );
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });
})