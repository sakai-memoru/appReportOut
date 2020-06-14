'use strict';
(function () {
    const fs = require('fs');
    const path = require('path');
    const { docopt } = require('docopt');
    const config = require('config');
    const cheerio = require('cheerio');
    const iconv = require('iconv-lite');
    // parameters in ./config/default.js
    let base_folder_param = '';
    let target_folder_param = '';
    let target_file_param = '';
    let output_folder_param = '';
    // - private method ---------------------------------------//
    let ensureDirSync = (dirpath) => {
        // ## renderTEmplate.ensureDirSync
        try {
            return fs.mkdirSync(dirpath);
        }
        catch (err) {
            if (err.code !== 'EEXIST')
                throw err;
        }
    };
    let transformHtml = (file_path) => {
        // ## renderTEmplate.mergeTemplate
        const CONS_MODULE_NAME = "renderTEmplate.mergeTemplate";
        let ret = '';
        try {
            // - get html contents
            let html_contents = fs.readFileSync(file_path); // FIXME how to read a file of shift_jis
            let buf = new Buffer(html_contents, 'binary'); // FIXME
            html_contents = iconv.decode(buf, "shift_jis");
            // console.log(html_contents);
            // - html parse
            let $ = cheerio.load(html_contents);
            //
            // - remove span tag
            $('td > span').each(function () {
                var childNodes = this.childNodes;
                $(this).replaceWith(childNodes);
                //$(this).removeAttr('style')
            });
            // - remove script and link tag
            $('script').each(function () {
                $(this).remove();
            });
            $('link').each(function () {
                $(this).remove();
            });
            // - append stylesheet.css link
            $('head').append('<link rel=Stylesheet href=stylesheet.css>');
            // - remove meta //FIXME
            $('meta').each(function () {
                $(this).remove();
            });
            $('head').append('<meta http-equiv="Content-Type" content="text/html; charset="UTF-8">');
            //
            ret = $('html').html();
            return ret;
        }
        catch (err) {
            throw CONS_MODULE_NAME + "\r\n" + err;
        }
    };
    // -----------------------------------------------------//
    // - main method ---------------------------------------//
    let doProcess = () => {
        // ## renderTemplate.doProcess
        const CONS_MODULE_NAME = "renderTemplate.doProcess";
        console.info('//------------------/ start !');
        // - generate pathes
        let target_folder = path.join(base_folder_param, target_folder_param);
        let target_file = path.join(target_folder, target_file_param);
        console.log(`target file = ${target_file}`);
        //
        let output_folder_def = path.join(base_folder_param, output_folder_param);
        let targetPathAry = path.dirname(target_folder).split(path.sep);
        let sub_folder = targetPathAry[targetPathAry.length - 1];
        console.log(`sub_folder = ${sub_folder}`);
        let output_folder = path.join(output_folder_def, sub_folder);
        ensureDirSync(output_folder);
        let output_file = path.join(output_folder, "templ.table.html");
        console.log(`output file = ${output_file}`);
        //
        try {
            // - get output_content
            let output_contents = transformHtml(target_file);
            // - copy a stylesheet
            fs.copyFileSync(path.join(target_folder, "Stylesheet.css"), path.join(output_folder, "Stylesheet.css"));
            // - output transformed html file
            fs.writeFileSync(output_file, output_contents, 'utf-8');
            console.info('//------------------/ end ...');
            return 0;
        }
        catch (err) {
            throw CONS_MODULE_NAME + "\r\n" + err;
        }
    };
    // ------------------------------------------// entry point
    if (typeof (WScript) !== 'undefined') {
        WScript.Echo('[Warn] Opps! Sorry. This app is for Nodejs.');
    }
    else {
        // ## renderTemplate Entry point
        // - define command line option statement
        const doc = `
    Usage: 
        index.js --run [--tempdir <temp_folder>] [--tempfile <temp_file>] [--output <output_folder> ]
        index.js -h | --help | --version
    `;
        let argv = docopt(doc, {
            "version": "0.0.1"
        });
        // - get config
        let CONF = config.get("CONF");
        // console.dir(CONF);
        //
        // - put opts into variables with config
        let flag = false;
        // put with command opts    
        if (argv["--run"]) {
            if (argv["--tempdir"]) {
                target_folder_param = argv["<temp_folder>"];
                flag = true;
            }
            else {
                target_folder_param = CONF['TARGET_FOLDER'];
            }
            ;
            if (argv["--tempfile"]) {
                target_file_param = argv["<temp_file>"];
            }
            else {
                target_file_param = CONF['TARGET_FILE'];
            }
            ;
            if (argv["--output"]) {
                output_folder_param = argv["<output_folder>"];
            }
            else {
                output_folder_param = CONF['OUTPUT_FOLDER'];
            }
            ;
            // console.dir(argv);
        }
        // - put with config
        if (CONF['BASE_FOLDER'] == '') {
            base_folder_param = __dirname;
        }
        else {
            base_folder_param = CONF['BASE_FOLDER'];
        }
        if (flag) {
            base_folder_param = '';
        }
        // ----------------------------------------------
        // ## do process
        try {
            process.exit(doProcess());
        }
        catch (err) {
            console.error("An Error has occurred!!");
            console.error(err);
        }
    }
}());
//# sourceMappingURL=index.js.map