# [Excel VBA] appReportOut : 帳票レイアウトからHTMLテンプレートを生成するツール

## overveiw :

- Excelで設計した帳票レイアウトから、HTMLを出力する設計補助ツール(VBA Excel)。HTMLでReport画面をプロトタイプとして作成する際に利用する。
- 帳票レイアウト定義書から、ひな形のHTMLテンプレートを生成することができる。定義情報は、JSONファイルにも出力できる。  
- 帳票HTMLテンプレート(report)に、JSONで設定したデータ値を、挿し込み、静的HTMLを生成する。

### 機能 :
- Excelにて、帳票レイアウトを作成し、指定フォルダ（input folder)に配置する。  
- メニューより、以下の機能を利用する。
  + 【Output JSON  Def】帳票レイアウトにて定義した情報をもとに、定義情報をJSONファイルに出力する。  
  + 【Generate HTML/CSS Template】帳票レイアウトにて定義した情報をもとに、プロトタイプ用の帳票テンプレート(HTML/CSS)を生成する。(CSS Grid形式)  
  + 【Dump Simple JSON Def】Layout確認のための定義情報を出力する。  
  + 【Generate Easy Template】帳票レイアウトより、Excel 標準でHTMLファイル保存によりできるHTMLをTemplateとして生成する。

- 【Batch機能】 : ./js/renderTemplate.js
    + Templateに、JSONで定義したデータ値を差し込み、静的なHTMLを生成する。

## Installation :

### Excel VBA tool

- GitHubより、Cloneする。  
    + https://github.com/sakai-memoru/appReportOut 

- 参照設定が必要。  

![image](https://gyazo.com/7d30f2387e7818067fd7596a82e507e9.png) 

### nodejs tool

- Excelで生成されたHTMLを一部加工のために、nodejsで作成したバッチアプリを利用している。実行には、nodejsがインストールされていること。以下、npmで、利用するmoduleをインストールする。  

```
$ npm install
$ node ./js/index.js
Usage:
        index.js --run [--tempdir <temp_folder>] [--tempfile <temp_file>] [--output <output_folder> ]
        index.js -h | --help | --version]
```
### 実行環境
- Microsoft Excel  
![excel](https://gyazo.com/0208cfb4b0e4c7e9494f35502677af34.png)  

- nodejs, typescript, powershell  
```
PS G:\Users\sakai> node --version
v13.8.0
PS G:\Users\sakai> npm --version
6.13.6
PS G:\Users\sakai> tsc --version
Version 3.9.5
PS G:\Users\sakai> $PSVersionTable

Name                           Value
----                           -----
PSVersion                      5.1.17134.858
PSEdition                      Desktop
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0...}
BuildVersion                   10.0.17134.858
CLRVersion                     4.0.30319.42000
WSManStackVersion              3.0
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
```  


## Usage :
- アプリは以下。
    - アプリ本体  ：appReportDef.xlsm  
        + Batch  
            - ProcessForReportSheet
  - アプリconfig：config.json  
  - 定義情報取得form：forms/\_\_TRANS_REPORTS__.xlsm  
      + GetDef
      + GetSimpleJson  
      + SaveAsHtmlTemplate 
  - 定義情報取得config：forms/\_\_TRANS_REPORTS__.config.json 

- appReporDef.xlsmを開く。Menuより起動する。 

![menu](https://gyazo.com/ddeefb0aaea9ff952dbcf095fda9d1ee.png)  

- 静的HTML生成バッチ(./js/rederTemplate.js)は、以下で起動。(起動時のConfigを、`./config/default.json`に設定する必要あり)

```
 
$ node ./js/renderTemplate.js --run
 
```


### 初期コンフィグ設定 :
   
```
{
    "BASE_FOLDER": "",
    "INPUT_FOLDER": "input",
    "OUTPUT_FOLDER": "output",
    "TEMP_FOLDER": "input/temp",
    "BACKUP_FOLDER": "input/backup",
    "FORM_FOLDER": "reports",
    "DATA_FOLDER": "data",
    "BACKUP_DATA_FOLDER": "data/backup",
    "REPORT_DEF": {
        "SHEET_TYPE": "REPORT",
        "INPUT_LIKE": "*.xlsx",
        "FORM_FILE": "__OUT_REPORT__.xlsm",
        "FORM_SHEET": "report",
        "DATA_LIKE": "*.json",
        "MACRO_GET_METHOD": "GetDef",
        "MACRO_OUT_METHOD": "OutputTemplate",
        "MACRO_OUTHTML_METHOD": "SaveAsHtmlTemplate",
        "MACRO_DUMP_METHOD": "DumpSimpleJson"
    },
    "CONTROL_PREFIX": "__",
    "SOURCE_FROM": "_source",
    "APP_NAME" : "appReportDef"
}
```
### Environment

![env](https://gyazo.com/e7635f7e49ef29455e5e1b88461da28c.png)


## Execution sample

- Excel Layout : input/RequestSheet.xlsx  

![layout](https://gyazo.com/0bc44a823a55c603575406402ba17025.png)  

- HTML Layout : output/RequestSheet_Sheet_200612210921/templ.table.html  

![template](https://gyazo.com/95754c949c8b2a8dda95e1f914d6ce09.png)

- HTML Rendered : data/RequestSheet_Sheet_200612210921/table0_001.html  

![rendered](https://gyazo.com/e7655f2d828d2cd16733a03fab52259f.png)

## application I/F
```vb
Public Function Batch( _
        ByVal datatype As String, _
        Optional ByVal outTemplOn As Variant = False, _
        Optional ByVal outTemplHtmlOn As Variant = False, _
        Optional ByVal dumpOn As Variant = False, _
        Optional ByVal moveOn As Variant = False _
    ) As Variant
'''' **********************************************
'' @function batch
'' @param datatype {String} 処理データタイプ
''        config.jsonのキー "REPORT_DEF"
'' @param outTemplOn {Variant<boolean>}
''            Template出力flag
'' @param outTemplHtmlOn {Variant<boolean>}
''            簡易Template出力flag
'' @param dumpOn {Variant<boolean>}
''            Dump JSON出力flag
'' @param moveOn  {Variant<boolean>}
''            Inputファイル移動flag
''
```
## note :
- 落ち着いたら、もう少し記述を追加します。  

## reference :

- 以下の外部ライブラリを使用しています。  
  + VBA-JSON : JsonConverter.bas  
    - https://github.com/VBA-tools/VBA-JSON  
  + MiniTemplator  
    - https://www.source-code.biz/MiniTemplator/  

// --- end of README.md