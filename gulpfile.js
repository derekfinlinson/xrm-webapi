/// <binding AfterBuild='default' Clean='clean' />

var gulp = require("gulp");
var ts = require("gulp-typescript");

gulp.task("default", function () {
    var tsResult = gulp.src("src/scripts/xrm-webapi.ts")
        .pipe(ts({
              noImplicitAny: true,
              out: "xrm-webapi.js"
        }));
    return tsResult.js.pipe(gulp.dest("dist"));
});