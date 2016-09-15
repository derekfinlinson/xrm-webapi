/// <binding AfterBuild='default' Clean='clean' />

var gulp = require("gulp");
var del = require("del");

gulp.task("clean", function () {
    return del(["dist/*"]);
});

gulp.task("default", 
    function () {
        gulp.src("src/xrm-webapi.ts").pipe(gulp.dest("dist"));
        gulp.src("typings/**/*").pipe(gulp.dest("dist/typings"));
    });