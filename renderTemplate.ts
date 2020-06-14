'use strict';

(function(){

const fs = require('fs');
const path = require('path');
const {docopt} = require('docopt');
const config = require('config');
const globby = require('globby');
const iconv = require('iconv-lite');
const jsonfile = require('jsonfile');
const Handlebars = require("handlebars");

// parameters in ./config/default.js
let base_folder_param :string = '';
let form_folder_param :string = '';
let form_file_param :string = '';
let data_folder_param :string = '';
let data_filelike_param :string = '';
let merged_folder_param :string = '';
let merged_file_base_param :string = '';

// - private method ---------------------------------------//
let ensureDirSync = (dirpath:string):any => {
    // ## renderTEmplate.ensureDirSync
    try {
      return fs.mkdirSync(dirpath)
    } catch (err) {
      if (err.code !== 'EEXIST') throw err
    }
}

let mergeTemplate = (templ_path:string, datafiles :string[]) :void => {
    // ## renderTEmplate.mergeTemplate
    const CONS_MODULE_NAME = "renderTEmplate.mergeTemplate";
    let ret:string = '';
    try{
        // - get paths
        let templPathAry :string[] = path.dirname(templ_path).split(path.sep);
        let sub_folder :string = templPathAry[templPathAry.length-1]
        console.log(`sub_folder = ${sub_folder}`);
        //
        let merged_folder_def:string = path.join(base_folder_param, merged_folder_param);
        let merged_folder:string = path.join(merged_folder_def, sub_folder);
        ensureDirSync(merged_folder);
        let merged_file_base:string = path.join(merged_folder,merged_file_base_param);
        // console.log(`merged file = ${merged_file}`);
        //
        // - get contents
        console.log(`templ_path = ${templ_path}`);
        let templ_form :string = fs.readFileSync(templ_path,"utf-8");
        //console.log(templ_form);
        //
        const template :any = Handlebars.compile(templ_form);
        //
        // - (())-loop in data files
        let merged_data :object[];
        let merged_str :string;
        let merged_file :string;
        datafiles.forEach((data_path:string, idx:number)=>{
            console.log(`data_path = ${data_path}`);
            merged_data = jsonfile.readFileSync(data_path,'utf-8');
            merged_file_base = merged_file_base + idx;
            // console.log(merged_data);
            merged_data.forEach((val, idx)=>{
                merged_str = template(val);
                //console.log(output_str)
                // - output transformed html file
                merged_file = merged_file_base + '_' + ('000' + idx).substr(-3) + ".html";
                fs.writeFileSync(merged_file, merged_str, 'utf-8');
                console.log(`merged_file = ${merged_file}`);
            });
        })
        // - copy a stylesheet
        fs.copyFileSync(path.join(path.dirname(templ_path),"Stylesheet.css"), path.join(merged_folder,"Stylesheet.css"))
        //
    } catch(err){
        throw CONS_MODULE_NAME + "\r\n" + err;
    }
}

// -----------------------------------------------------//
// - main method ---------------------------------------//
let doProcess =  () :any => {
    // ## renderTemplate.doProcess
    const CONS_MODULE_NAME = "renderTemplate.doProcess"
    console.info('//------------------/ start !')
    // - generate pathes
    let form_folder:string = path.join(base_folder_param,form_folder_param);
    let form_file:string = path.join(form_folder,form_file_param);
    console.log(`form_file = ${form_file}`);
    //
    let data_folder:string = path.join(base_folder_param, data_folder_param);
    // console.log(`data folder = ${data_folder}`);
    let data_filelike = path.join(data_folder,data_filelike_param)
    data_filelike = data_filelike.split('\\').join('//'); //FIXME
    //
    try {
        // - get files with glob
        let aryfiles :string[] = globby.sync(data_filelike);
        //let aryfiles: string[] = globby.sync("g://Users//sakai//Desktop//ExcelVbaApp//ReportOutApp//data//*.json");
        // console.dir(data_filelike);
        // console.dir(aryfiles);
        //
        // ## merging template func "mergeTemplate"
        mergeTemplate(form_file,aryfiles);
        console.info('//------------------/ end ...')
        return 0;    
    } catch(err) {
        throw  CONS_MODULE_NAME + "\r\n" + err;
    }
}

// ------------------------------------------// entry point
if(typeof(WScript) !== 'undefined'){
    WScript.Echo('[Warn] Opps! Sorry. This app is for Nodejs.');
} else {
    // ## renderTemplate Entry point
    // - define command line option statement
    const doc = `
    Usage: 
        index.js --run [--formdir <form_folder>] [--formfile <form_file>] [--mergeddir <merged_folder> ] [--datadir <data_folder>]
        index.js -h | --help | --version
    `
    let argv:object = docopt(doc, {
        "version": "0.0.1"
    });
    // - get config
    let CONF: object = config.get("CONF");
    // console.dir(CONF);
    //
    // - put opts into variables with config
    let flag :boolean = false;
    if (argv["--run"]) {
        if (argv["--formdir"]) {
            form_folder_param = argv["<form_folder>"];
            flag = true;
        }else{
            let target_folder_param = CONF['TARGET_FOLDER'];
            let aryTargetPath = target_folder_param.split('//');
            let target_sub_folder = aryTargetPath[aryTargetPath.length-2];
            console.log(`target_sub_folder = ${target_sub_folder}`);
            form_folder_param = path.join(CONF['OUTPUT_FOLDER'], target_sub_folder);
        };
        if (argv["--tempfile"]) {
            form_file_param = argv["<form_file>"];
        }else{
            form_file_param = CONF['OUTPUT_FILE'];
        };    
        if (argv["--datadir"]) {
            data_folder_param = argv["<data_folder>"];
        }else{
            data_folder_param = CONF['DATA_FOLDER'];
        };    
        if (argv["--mergeddir"]) {
            merged_folder_param = argv["<merged_folder>"];
        }else{
            merged_folder_param = CONF['MERGED_FOLDER'];
        };    
        // console.dir(argv);
    }
    // - put with config
    if (CONF['BASE_FOLDER'] == '' ) {
        base_folder_param = __dirname;
    } else {
        base_folder_param = CONF['BASE_FOLDER'];
    }
    if (flag){
        base_folder_param = '';
    }
    data_filelike_param = CONF['DATA_FILELIKE'];
    merged_file_base_param = CONF['MERGED_FILE_BASE'];
    // ----------------------------------------------
    // ## do process
    try {
        process.exit(doProcess());
    } catch(err) {
        console.error("An Error has occurred!!")
        console.error(err);
    }
}

}())
