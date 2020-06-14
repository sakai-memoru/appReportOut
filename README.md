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
  + 【Dump Simple JSON Def】帳票レイアウトに、外部ファイル(JSON)で定義した入力データを挿し込み、PDF(Excel)を生成する。  
  + 【Dump Def】Layout確認のようための定義情報を出力する。

- 【Batch機能】 : renderTemplate.js
    + Templateに、JSONで定義したデータ値を差し込み、静的なHTMLを生成する。

## Installation :

- GitHubより、Cloneする。  

- 参照設定が必要。  

![image](https://gyazo.com/7d30f2387e7818067fd7596a82e507e9.png) 

- Excelで生成されたHTMLを一部加工のために、nodejsで作成したバッチアプリを利用している。実行には、nodejsがインストールされていること。以下、npmで、利用するmoduleをインストールする。  

```
$ npm install
$ node ./js/index.js
Usage:
        index.js --run [--tempdir <temp_folder>] [--tempfile <temp_file>] [--output <output_folder> ]
        index.js -h | --help | --version]
```

## Usage :
- アプリは以下。
  - アプリ本体  ：appReportDef.xlsm  
    + Batch  
      - ProcessForReportSheet
  - アプリconfig：config.json  
  - 定義情報取得form：forms/\_\_TRANS_REPORTS__.xlsm  
    + GetReportDef
    + GetDefDump  
  - 定義情報取得config：forms/\_\_TRANS_REPORTS__.config.json 

- appReporDef.xlsmを開く。Menuより起動する。 

![menu](https://gyazo.com/ddeefb0aaea9ff952dbcf095fda9d1ee.png)  

- 静的HTML生成バッチは、以下で起動。(起動時のConfigを、`./config/default.json`に設定する必要あり)

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

![env](https://gyazo.com/77fcdd24660acfc1b477ab985861c2a7.png)


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