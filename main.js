// Compiled using ts2gas 3.6.2 (TypeScript 3.9.3)
/****************************************************************************************************

# officeTranslator

    copy > conversion > translation > reconversion > mail > move

## using
    Advanced google service : drive api

## next tasks
    - multi attach file
    - create in folder
****************************************************************************************************/

function Translator() {
    //get configration spreadsheet 
    var st = SpreadsheetApp.getActive().getSheetByName("main");
    var params = st.getDataRange().getValues();
    //get id source/destination folder and more parameter
    var sourceFolder = DriveApp.getFolderById(params[2][1]);
    var destinationFolderId = params[3][1];
    var destinationFolder = DriveApp.getFolderById(destinationFolderId);
    var sourceLanguage = params[4][1];
    var translateLanguage = params[5][1];
    var mailenable = params[6][1];
    var mailAddress = params[7][1];
    var mailsubject = params[8][1];
    var translationType = params[9][1]; // for word and powerpoint only
    var fontcolor = params[10][1]; // for word and powerpoint only
    var multisheet = params[11][1];
    var files = sourceFolder.getFiles();
    var workingFile; //
    var sourceFile; //
    var doc;
    var sheets;
    var mimetype;
    var mimecode;
    var workingFileId;
    var resultPdf;
    var sourceFileName;

    //Check each file in the sourceFolder
    while (files.hasNext()) {
        sourceFile = files.next();
        sourceFileName = sourceFile.getName();
        mimetype = sourceFile.getMimeType();
      
        if (sourceFileName === "OfficeTranslator"){
            continue;
        }
        // Branch processing by mime type
        switch (mimetype) {
            
            //Google spread sheets
            case "application/vnd.google-apps.spreadsheet":
                workingFile = sourceFile.makeCopy("翻訳" + sourceFileName, destinationFolder);
                workingFileId = workingFile.getId();
                if (multisheet === true){
                    sheets = SpreadsheetApp.openById(workingFileId).getSheets();
                    sheets.forEach( sheet => {
                        doc = sheet;
                        SpsTranslator(doc, sourceLanguage, translateLanguage);
                    });
                    sheetId = "";
                }else{
                    doc = SpreadsheetApp.openById(workingFileId).getSheets()[0];
                    sheetId = doc.getSheetId();
                    SpsTranslator(doc, sourceLanguage, translateLanguage);
                    
                }
                SpreadsheetApp.flush();
                resultPdf = ss2pdf(doc, workingFileId, sheetId, destinationFolder,multisheet,sourceFileName);
                break;

            //Google Documents
            case "application/vnd.google-apps.document":
                workingFile = sourceFile.makeCopy("翻訳" + sourceFileName, destinationFolder);
                workingFileId = workingFile.getId();
                doc = DocumentApp.openById(workingFileId);
                docTranslator(doc, sourceLanguage, translateLanguage, translationType, fontcolor);
                resultPdf = doc2pdf(doc, workingFileId, destinationFolder, mimetype);
                break;

            //Google Slides
            case "application/vnd.google-apps.presentation":
                workingFile = sourceFile.makeCopy("翻訳" + sourceFileName, destinationFolder);
                workingFileId = workingFile.getId();
                doc = SlidesApp.openById(workingFileId);
                ppTranslator(doc, sourceLanguage, translateLanguage, translationType, fontcolor);
                resultPdf = doc2pdf(doc, workingFileId, destinationFolder, mimetype);
                break;

            //Excel
            case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                mimecode = "GOOGLE_SHEETS";
                options = {
                    title: sourceFileName,
                    mimeType: MimeType[mimecode],
                    parents: [{ id: destinationFolderId }]
                };
                workingFile = Drive.Files.insert(options, sourceFile.getBlob());
                workingFileId = workingFile.getId();
                doc = SpreadsheetApp.openById(workingFileId).getSheets()[0];
                sheetId = doc.getSheetId();
                SpsTranslator(doc, sourceLanguage, translateLanguage);
                google2ms(workingFileId, destinationFolder, mimetype);
                resultPdf = ss2pdf(doc, workingFileId, sheetId, destinationFolder,multisheet,sourceFileName);
                break;

            //powerpoint
            case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
                mimecode = "GOOGLE_SLIDES";
                options = {
                    title: sourceFileName,
                    mimeType: MimeType[mimecode],
                    parents: [{ id: destinationFolderId }]
                };
                workingFile = Drive.Files.insert(options, sourceFile.getBlob());
                workingFileId = workingFile.getId();
                doc = SlidesApp.openById(workingFileId);
                ppTranslator(doc, sourceLanguage, translateLanguage, translationType, fontcolor);
                google2ms(workingFileId, destinationFolder, mimetype);
                resultPdf = doc2pdf(doc, workingFileId, destinationFolder, mimetype);
                break;


            //word
            case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                mimecode = "GOOGLE_DOCS";
                options = {
                    title: sourceFileName,
                    mimeType: MimeType[mimecode],
                    parents: [{ id: destinationFolderId }]
                };
                workingFile = Drive.Files.insert(options, sourceFile.getBlob());
                workingFileId = workingFile.getId();
                doc = DocumentApp.openById(workingFileId);
                docTranslator(doc, sourceLanguage, translateLanguage, translationType, fontcolor);
                google2ms(workingFileId, destinationFolder, mimetype);
                resultPdf = doc2pdf(doc, workingFileId, destinationFolder, mimetype);
                break;


            default:
                break;
        }
        destinationFolder.addFile(sourceFile); //移動先のフォルダに先にファイルを追加
        sourceFolder.removeFile(sourceFile); //ソースファイルを削除
    }
    if (mailenable === true) {
        sendmail(mailAddress, resultPdf, mailsubject);
    }
    ;
}

/*
translation word 
*/
function docTranslator(doc, sourceLanguage, translateLanguage, translationType, fontcolor) {
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    var textBefore; // 翻訳前テキスト
    var textTranslated; // 翻訳後テキスト
    var mycolor = color[fontcolor];
    if (mycolor == "") {
        mycolor = color.lightgray;
    }
    paragraphs.forEach( p => {
        textBefore = p.getText();
        if (textBefore != "") {
            // 翻訳テキストを原文と置換するタイプ
            if (translationType === "テキスト置換") {
                textTranslated = LanguageApp.translate(textBefore, language[sourceLanguage], language[translateLanguage]);
                p.setText(textTranslated);
            }
            else { //翻訳テキストを原文の後ろに追記タイプ
                var text = p.editAsText();
                var textLength = text.getText().length;
                textTranslated = LanguageApp.translate(textBefore, language[sourceLanguage], language[translateLanguage]);
                p.appendText(" ");
                p.appendText(textTranslated);
                text.setForegroundColor(textLength, textLength + textTranslated.length, mycolor); //gray
            };
        }
    });
    doc.saveAndClose();
}
/*
translation ppt
*/
function ppTranslator(presentation, sourceLanguage, translateLanguage, translationType, fontcolor) {
    var slides = presentation.getSlides();
    var mycolor = color[fontcolor];
    if (mycolor == "") {
        mycolor = color.lightgray;
    }
    //Replacement
    slides.forEach(slide => {
        var shapes = slide.getShapes();
        shapes.forEach(shape => {
            var rangeA = shape.getText();
            var textBefore = rangeA.asString();
            var textTranslated = LanguageApp.translate(textBefore, language[sourceLanguage], language[translateLanguage]);
            if (translationType === "テキスト置換") {
                rangeA.setText(textTranslated);
            }
            else {
                rangeA.appendParagraph(""); //改行してるだけ
                var rangeB = rangeA.appendText(textTranslated);
                var style = rangeB.getTextStyle();
                var fontsize = style.getFontSize();
                style.setForegroundColor(mycolor) //グレー
                    .setFontSize(fontsize / 2); //翻訳前テキストの半分のフォントサイズ
            }
        });
    });
    presentation.saveAndClose();
}
/*
translation spreadsheet
*/
function SpsTranslator(targetSheet, sourceLanguage, translateLanguage) {
    var maxrow = targetSheet.getLastRow();
    var maxcol = targetSheet.getLastColumn();
    var text;
    for (var j = 1; j <= maxcol; j++) {
        for (var i = 1; i <= maxrow; i++) {
            text = targetSheet.getRange(i, j).getValue();
            if (text != "" && typeof (text) === "string") {
                var sourceDoc = targetSheet.getRange(i, j).getValue();
                var translate = LanguageApp.translate(sourceDoc, language[sourceLanguage], language[translateLanguage]);
                targetSheet.getRange(i, j).setValue(translate);
            }
        }
    }
}

/*
 GSpreadsheet > MsExcel
 Gdocument > MsWord
*/
function google2ms(id, folder, mimetype) {
    var new_file;
    var url;
    var doc;
    var extension;
    var filename;
    switch (mimetype) {
        //ソースファイルがワード
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            url = `https://docs.google.com/document/d/${id}/export?format=docx` ;
            doc = DocumentApp.openById(id);
            extension = "docx";
            break;
        //ソースファイルがパワポ
        case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
            url = `https://docs.google.com/presentation/d/"${id}"/export/pptx`;
            doc = SlidesApp.openById(id);
            extension = "pptx";
            break;
        //ソースファイルがスプシ
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            url = `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`;
            doc = SpreadsheetApp.openById(id);
            extension = "xlsx";
            break;
    }
    var options = {
        method: "get",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() == 200) {
        filename = doc.getName();
        new_file = folder.createFile(res.getBlob()).setName("翻訳済" + filename + "." + extension);
    }
    return new_file;
}
/*
 GSpreadSheet > pdf
*/
function ss2pdf(doc, spreadSheetId, targetSheetId, folder, multisheet,fileName) {
    var new_file;
    var url;
    if (multisheet === true) {
        url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/export?exportFormat=pdf`;
    }
    else {
        url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/export?exportFormat=pdf&gid=SID`.replace("SID", targetSheetId);
    }
    var options = {
        method: "get",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() == 200) {
        new_file = folder.createFile(res.getBlob()).setName("翻訳" + fileName + ".pdf");
    }
    return new_file;
}
/*
 G slide > pdf
 G document > pdf
*/
function doc2pdf(doc, id, folder, mimetype) {
    var new_file;
    var url;
    var filename;
    switch (mimetype) {

        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document": 
        case "application/vnd.google-apps.document"://G Docs
            url = `https://docs.google.com/document/d/${id}/export?exportFormat=pdf`;
            break;

        case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        case "application/vnd.google-apps.presentation": //S slide
            url = `https://docs.google.com/presentation/d/${id}/export/pdf`;
            break;
    }
    var options = {
        method: "get",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() == 200) {
        filename = doc.getName();
        new_file = folder.createFile(res.getBlob()).setName(filename + ".pdf");
    }
    return new_file;
}
/*
 メールを送信する関数
*/
function sendmail(address, file, mailsubject) {
    var bodymsg1 = "こんにちは。\nこのメールはGoogle Appsのプログラムから自動送信で送信しています。";
    var bodymsg2 = "Hello,dear.\nThis email is automatically sent from Google Apps.";
    GmailApp.sendEmail(address, mailsubject, bodymsg1 + "\n\n" + bodymsg2, { attachments: [file] });
}
var language = {
    "クメール語": "km",
    "キニヤルワンダ語": "rw",
    "ノルウェー語": "no",
    "アラビア文字": "ar",
    "スンダ語": "su",
    "ミャンマー語（ビルマ語）": "my",
    "リトアニア語": "lt",
    "エストニア語": "et",
    "ベラルーシ語": "be",
    "ブルガリア語": "bg",
    "アフリカーンス語": "af",
    "マルタ語": "mt",
    "タタール語": "tt",
    "フランス語": "fr",
    "マレー語": "ms",
    "ポルトガル語（ポルトガル、ブラジル）": "pt",
    "イディッシュ語": "yi",
    "アイルランド語": "ga",
    "モンゴル語": "mn",
    "セブ語": "ceb",
    "サモア語": "sm",
    "カンナダ語": "kn",
    "ボスニア語": "bs",
    "ラテン語": "la",
    "タミル語": "ta",
    "マラヤーラム文字": "ml",
    "オリヤ語": "or",
    "アムハラ語": "am",
    "マケドニア語": "mk",
    "スペイン語": "es",
    "クロアチア語": "hr",
    "インドネシア語": "id",
    "パンジャブ語": "pa",
    "ネパール語": "ne",
    "ショナ語": "sn",
    "エスペラント語": "eo",
    "パシュト語": "ps",
    "アイスランド語": "is",
    "モン語": "hmn",
    "マラガシ語": "mg",
    "タイ語": "th",
    "ヨルバ語": "yo",
    "フィンランド語": "fi",
    "チェコ語": "cs",
    "アルメニア語": "hy",
    "マオリ語": "mi",
    "フリジア語": "fy",
    "ヒンディー語": "hi",
    "ウクライナ語": "uk",
    "トルコ語": "tr",
    "ロシア語": "ru",
    "ベトナム語": "vi",
    "シンハラ語": "si",
    "テルグ語": "te",
    "ポーランド語": "pl",
    "ペルシャ語": "fa",
    "セソト語": "st",
    "タガログ語（フィリピン語）": "tl",
    "ウルドゥー語": "ur",
    "ウイグル語": "ug",
    "アゼルバイジャン語": "az",
    "セルビア語": "sr",
    "イボ語": "ig",
    "ルーマニア語": "ro",
    "スウェーデン語": "sv",
    "ヘブライ語": "he",
    "ラトビア語": "lv",
    "カザフ語": "kk",
    "スワヒリ語": "sw",
    "日本語": "ja",
    "デンマーク語": "da",
    "コルシカ語": "co",
    "ラオ語": "lo",
    "タジク語": "tg",
    "コーサ語": "xh",
    "韓国語": "ko",
    "オランダ語": "nl",
    "ハワイ語": "haw",
    "ルクセンブルク語": "lb",
    "スコットランド ゲール語": "gd",
    "ズールー語": "zu",
    "ガリシア語": "gl",
    "ベンガル文字": "bn",
    "シンド語": "sd",
    "キルギス語": "ky",
    "ハウサ語": "ha",
    "グジャラト語": "gu",
    "英語": "en",
    "クルド語": "ku",
    "ドイツ語": "de",
    "中国語（繁体）": "zh-TW",
    "バスク語": "eu",
    "クレオール語（ハイチ）": "ht",
    "ソマリ語": "so",
    "スロベニア語": "sl",
    "トルクメン語": "tk",
    "グルジア語": "ka",
    "ジャワ語": "jv",
    "カタロニア語": "ca",
    "イタリア語": "it",
    "ウェールズ語": "cy",
    "スロバキア語": "sk",
    "ウズベク語": "uz",
    "アルバニア語": "sq",
    "中国語（簡体）": "zh-CN",
    "ハンガリー語": "hu",
    "ギリシャ語": "el",
    "マラーティー語": "mr",
    "ニャンジャ語（チェワ語）": "ny"
};
var color = {
    "変更しない": "",
    "lightgray": "#d3d3d3",
    "black": "#000000",
    "red": "#ff0000",
    "royalblue": "#4169e1",
    "white": "#ffffff"
};
