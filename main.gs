//経路案内
function getMapUrl(starting, destination, _by) {
  
  if(_by == "public"){

    by = "data=!4m2!4m1!3e3"; //公共交通機関
  }
  else if(_by == "car"){

    by = "data=!4m3!4m2!3e0!4e1"; //車
  }
  else if(_by == "bus"){

    by = "data=!4m4!4m3!2m2!4e2!5e0"; //バス
  }
  else{

    by = "";
  }

  
  var mapUrl = `https://www.google.co.jp/maps/dir/${starting}/${destination}/${by}`;

  return mapUrl;
}

//スプレッドシートからAIのキャラクター設定を取得する
function getAiSettings(){
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aiSettings = ss.getRangeByName("ai_settings").getValue();
    return aiSettings;
  } catch(e){
    return null;
  }
}

//promptの内容をキャラクター設定されたAI（ChatGPT API）に投げて回答を取得する
function getReply(count, prompt_u, prompt_a){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusSheet = ss.getSheetByName("status");
  statusSheet.getRange("B7").setValue("in GetReply OK");
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  let messages = [];
  //以前の会話をプロンプトとして入力
  for(var i=0; i<count; i++){
    messages.push({
      "role": "user",
      "content": prompt_u[i],
    });
    messages.push({
      "role": "assistant",
      "content": prompt_a[i],
    })
  }
  //今回の呼び出しのユーザーの入力を最後に追加
  messages.push({
    "role": "user",
    "content": prompt_u[prompt_u.length - 1],
    });

  const characterSettings = getAiSettings();
  
  if(characterSettings != null){
    messages.unshift({
      "role": "system",
      "content": characterSettings
    });
  }

  const payload = {
    "model": "gpt-3.5-turbo",
    "temperature" : 0.5, //0〜1で設定。大きいほどランダム性が強い
    "max_tokens": 500, //LINEのメッセージ文字数制限が500文字なので、それに合わせて調整
    "messages": messages
  };

  const requestOptions = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer "+ OPENAI_API_KEY
    },
    "payload": JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);

  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const retObj = {
    "payload": payload,
    "message": json.choices[0].message.content,
    "usage": json.usage
  };
  return retObj;
}

//file（音声ファイル）の内容をWhisperAPIを利用して文字列として取得する
function speechToText(file){
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  const payload = {
    "model": "whisper-1",
    "temperature" : 0,
    "language": "ja", //日本語以外にも対応する場合はこのプロパティは外す
    "file": file
  };
  
  const requestOptions = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer "+ OPENAI_API_KEY
    },
    "payload": payload
  };
  try{
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", requestOptions);

    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    const text = json.text
    return text;
  } catch(e){
    return e.message;
  }
}

// GPT-4V APIを呼び出す関数
function idenPicture(base64Image) {
  const PROMPT = "300文字以内でこの画像に写っているものや、撮られたであろう場所を列挙してください。";
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  let url = 'https://api.openai.com/v1/chat/completions';
  let options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + OPENAI_API_KEY,
    },
    payload: JSON.stringify({
      'model': 'gpt-4-vision-preview',
      'messages': [
        {
          'role': 'user',
          'content': [
            {
              'type': 'text',
              'text': PROMPT
            },
            {
              'type': 'image_url',
              'image_url': {
                'url': "data:image/jpeg;base64," + base64Image
              }
            }
          ]
        }
      ],
      'max_tokens': 200 // 返信の最大トークン数
    })
  };
  let response = UrlFetchApp.fetch(url, options);
  let responseMessage = JSON.parse(response).choices[0].message.content;
  return responseMessage;
}

function getCreateImg(prompt){
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  let url = 'https://api.openai.com/v1/images/generations';
  let options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + OPENAI_API_KEY,
    },
    payload: JSON.stringify({
      'model': "dall-e-3",
      'prompt': prompt,
      'n': 1,
      'size': "1024x1024"
    })
  }
  let response = UrlFetchApp.fetch(url, options);
  let img_url = JSON.parse(response).data[0].url;
  return img_url
}

//LINEでユーザーから送られてきた音声ファイルを取得する
function getContentByUser(messageId){
  const LINE_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  const requestOptions = {
    'headers': {
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
    },
    'method': 'get'
  };
  const response = UrlFetchApp.fetch(url, requestOptions);
  return response.getBlob().setName(`${messageId}.m4a`); //拡張子を指定しないとWhisperAPI側でエラーになるので注意
}

//LINEでユーザーから送られてきたメッセージ（スタンプや音声含む）を文字列に変換する
function convertMessageObjToText(messageObj){
  let result = "";
  const messageType = messageObj.type;
  switch(messageType){
    case "text": //文字列
      result = messageObj.text;
      break;
    case "sticker": //スタンプ。キーワードが設定されていればそれを取得する
      if(messageObj.keywords === undefined){
        result = "？？？";
      } else{
        result = messageObj.keywords.join(",");
      }
      break;
    case "image": //画像
      result = "この画像が分かりますか？";
      break;
    case "video": //動画
      result = "この動画が分かりますか？";
      break;
    case "audio": //音声。文字起こしする
      if(messageObj.contentProvider.type === "line"){
        const audioFile = getContentByUser(messageObj.id);
        const transcriptedText = speechToText(audioFile);
        result = transcriptedText;
      } else{
        result = "この音声が聞こえますか？";
      }
      break;
    case "file": //ファイル
      result = "このファイルは見られますか？";
      break;
    case "location": //位置情報
      let locationInfo = messageObj.title ? messageObj.title + "\n" : "";
      locationInfo += messageObj.address ? messageObj.address + "\n" : "";
      locationInfo += `latitude:${messageObj.latitude}\n`;
      locationInfo += `longitude:${messageObj.longitude}`;
      result = "ここはどんな場所ですか？\n" + locationInfo;
      break;
    default: //その他
      result = "？？？";
  }
  return result;
}

//logシートにログを出力する
function appendLog(logArray){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("log");
  logSheet.appendRow(logArray);
}

//statusシートにログを出力する
function appendLog_status(userRow, count, userMessage, botMessage){
  const alp = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q","R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

  const userID_column = alp[0]; //A
  const lang_column = alp[1]; //B
  const aiAvailablekey_column = alp[2]; //C
  const destination_column = alp[3]; //D
  const stampCount_column = alp[4]; //E
  const stampCount_n = 4;
  const logCount_column = alp[10]; //K
  const plomptStart_column = alp[11] //L
  const plomptStart_n = 11;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusSheet = ss.getSheetByName("status");
  const userStatusSheet = ss.getSheetByName("userStatus");
  statusSheet.getRange("B9").setValue("in appendLog_status OK");

  userStatusSheet.getRange(`${alp[plomptStart_n+2*count]}${userRow}`).setValue(userMessage);
  userStatusSheet.getRange(`${alp[plomptStart_n+2*count+1]}${userRow}`).setValue(botMessage);
}

//Routeシートからあらかじめ用意したルートを変えす(現在は使わなそう)
function findRoute(sheet, val){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(var i=1;i<dat.length;i++){
    if(dat[i][0] === val){
      return dat[i][1];
    }
  }
  return "error";
}

//userIdをもとに行番号を返す
function findUserRow(userId){
  //シートの取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusSheet = ss.getSheetByName("status");
  const userStatusSheet = ss.getSheetByName("userStatus");
  statusSheet.getRange("F4").setValue("get Sheet OK");

  //シートの全データを二次元配列で取得
  var dat = userStatusSheet.getDataRange().getValues();
  statusSheet.getRange("F5").setValue("get values OK");
  
  //検索開始
  for(var i=1; i<dat.length;i++){
    if(dat[i][0] == userId){
      statusSheet.getRange("F6").setValue("return OK");
      return i + 1;
    }
  }
  return -1; //userStatusシートにuserIdが登録されていない
}

//userStatusを初期情報を入れる
//友達を追加したときにこの関数が走る予定
function setUserStatus_initial(userId){
  //シート取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userStatusSheet = ss.getSheetByName("userStatus");

  //statusシートのfollowersをインクリメント
  const statusSheet = ss.getSheetByName("status");
  statusSheet.getRange("F2").setValue(Number(statusSheet.getRange("F2").getValue()) + 1);

  //userStatusSheet.appendRow([userID, lang, AI_available_key, destination, stamp_count, stamp_1, stamp_2, stamp_3, stamp_4, stamp_5, log_count, prompt])の初期値設定
  userStatusSheet.appendRow([userId, "ja", "false", "", 0, "", "", "", "", "",  0, ""]); //行追加
}

//LINE側へ送信するデータを用意する(テキスト)
function setRequestOptions(textdata, replyToken){
  const LINE_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
  var requestOptions = {
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
      },
      'method': 'post',
      'payload': JSON.stringify({
        'replyToken': replyToken,
        'messages': [{
          'type': 'text',
          'text': textdata,
        }]
      })
    };

  return requestOptions;
}

function setColumns(stamp_count, userRow){
  //シート取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userStatusSheet = ss.getSheetByName("userStatus");

  const alp = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q","R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

  const stampCount_n = 4;

  var columns = [];
  for(var i=1; i<=stamp_count; i++){
    columns.push({
      "imageUrl": userStatusSheet.getRange(`${alp[stampCount_n+i]}${userRow}`).getValue(),
      "action": {
        "type": "postback",
        "label": `${i}`,
        "data": `push ${i}`
      }
    })
  }

  return columns;
}

//正常に処理が終了したことを示すオブジェクトを生成&returnする
function returnSuccess(){
  return ContentService.createTextOutput(JSON.stringify({"content": "success"})).setMimeType(ContentService.MimeType.JSON);
}

function test(){
  const LINE_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const imageSheet = ss.getSheetByName("image");
  /*const requestOptions = {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
    },
    'method': 'get',
  }*/
  /*const requestOptions = {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      //"muteHttpExceptions" : true,
      "size": {
        "width": 1200,
        "height": 405
      },
      "selected": true,
      "name": "rich menu en",
      "chatBarText": "Menu",
      "areas": [
        {
          "bounds": {
            "x": 0,
            "y": 0,
            "width": 400,
            "height": 405
          },
          "action": {
            "type": "message",
            "label": "Search",
            "text": "Search"
          }
        },
        {
          "bounds": {
            "x": 401,
            "y": 0,
            "width": 400,
            "height": 405
          },
          "action": {
            "type": "message",
            "label": "Go",
            "text": "Go"
          }
        },
        {
          "bounds": {
            "x": 801,
            "y": 0,
            "width": 400,
            "height": 405
          },
          "action": {
            "type": "uri",
            "label": "Photo",
            "uri": "https://line.me/R/nv/camera/"
          }
        }
    ]
    })
  }
  //UrlFetchApp.fetch("https://api.line.me/v2/bot/richmenu/validate", requestOptions).getContent();
  UrlFetchApp.fetch("https://api.line.me/v2/bot/richmenu", requestOptions);*/
  const requestOptions = {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
    },
    'method': 'get',
  }
  json = UrlFetchApp.fetch("https://api.line.me/v2/bot/richmenu/list", requestOptions).getContentText();
  console.log(json);
}

//リクエストが送られるとこの関数が実行される
function doPost(e) {
  const LINE_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
  const events = JSON.parse(e.postData.contents).events;
  const url = 'https://api.line.me/v2/bot/message/reply';
  const alp = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q","R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
  const stampsLimit = 3; //スタンプ上限
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusSheet = ss.getSheetByName("status");
  const routeSheet = ss.getSheetByName("route");
  const userStatusSheet = ss.getSheetByName("userStatus");
  const richMenuSheet = ss.getSheetByName("richMenu");

  //userStatusシートの各カラムを設定
  const userID_column = alp[0]; //A
  const lang_column = alp[1]; //B
  const aiAvailablekey_column = alp[2]; //C
  const destination_column = alp[3]; //D
  const stampCount_column = alp[4]; //E
  const stampCount_n = 4;
  const logCount_column = alp[10]; //K
  const plomptStart_column = alp[11] //L
  const plomptStart_n = 11;

  //richMenuId
  const richMenuId_ja = "richmenu-bcce351994db80d479be5d067369f43a";
  const richMenuId_en = "richmenu-4e84432a2148ea0df5d5cc097afdb173";

  statusSheet.getRange("H1").setValue("sheet OK");

  const event = events[0];
  const replyToken = event.replyToken;
  //id取得
  const userID = event.source.userId;
  let userRow = findUserRow(userID);
  if(userRow == -1){ //userStatusシートに追加
    setUserStatus_initial(userID);
    userRow = findUserRow(userID);
  }
  const log_count = Number(userStatusSheet.getRange(`${logCount_column}${userRow}`).getValue()); //会話が何回目か

  //言語取得
  const userLang = userStatusSheet.getRange(`${lang_column}${userRow}`).getValue();

  statusSheet.getRange("H2").setValue("const OK");

  if(event.type == "memberJoined"){ //友達追加のとき
    if(findUserRow(userID) == -1)setUserStatus_initial(userID); //userStatusシートに送信元のuserIDがない場合は初期化
    return returnSuccess();
  }else if(event.type == "unfollow"){ //ブロックされたとき
    statusSheet.getRange("F2").setValue(Number(statusSheet.getRange("F2").getValue()) - 1); //ユーザー数をデクリメント
    return returnSuccess();
  }
  else if(event.type == "postback"){ //ポストバックイベントのとき
    var requestOptions = setRequestOptions(event.postback.data, replyToken);
    UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
    return returnSuccess();
  }

  const userMessage = convertMessageObjToText(event.message);

  statusSheet.getRange("B4:B10").clearContent() // ~ OKのログを削除

  if(event.message.type == "image"){

    // 画像が送られた時
    statusSheet.getRange("J1").setValue("in if image OK");
    let messageId = event.message.id;
    statusSheet.getRange("J2").setValue("get messageId OK");
    let imgUrl = "https://api-data.line.me/v2/bot/message/" + messageId + "/content"; //messageIdから画像取得用URLの作成
    
    statusSheet.getRange("J3").setValue("get imgUrl OK");

    // 送信された画像取得
    let image = UrlFetchApp.fetch(imgUrl, { 
      "headers": {
        "Content-Type": "application/json; charset=UTF-8",
        "Authorization": "Bearer " + LINE_ACCESS_TOKEN,
      }, 
      "method": "get"
    });

    let imageData = image.getContent();
    statusSheet.getRange("J4").setValue("get image.getContent() OK");
    let base64Image = Utilities.base64Encode(imageData); // 取得した画像をBase64エンコード
    statusSheet.getRange("J5").setValue("get base64Image OK");
    var botAnswer = idenPicture(base64Image); // OpenAI APIを呼ぶ関数(エンコード済み画像を渡す)
    statusSheet.getRange("J6").setValue("get botAnswer OK");

    let _destination = userStatusSheet.getRange(`${destination_column}${userRow}`).getValue();
    if(botAnswer.includes(_destination)){
      let stampCount = Number(userStatusSheet.getRange(`${stampCount_column}${userRow}`).getValue()) + 1;
      let matchedMessage;
      if(userLang == "ja"){
        matchedMessage = "スタンプ一覧は「スタンプラリー」と送信でみることができます！";
      }else if(userLang == "en"){
        matchedMessage = "The list of stamps can be viewed by clicking on Stamp Rally and sending a message!";
      }
      let numberOfStamps = 0;
      if(userStatusSheet.getRange(`${stampCount_column}${userRow}`).getValue() >= stampsLimit-1){

        numberOfStamps = stampCount;
        stampCount = 0;

        if(userLang == "ja"){
          matchedMessage = `スタンプが ${numberOfStamps} 個貯まりましたのでクーポンを贈呈します！\nhttps://lin.ee/QFd6bnz`;
        }else if(userLang == "en"){
          matchedMessage = `You have accumulated ${numberOfStamps} stamps and will receive a coupon!\nhttps://lin.ee/QFd6bnz`;
        }

      }
      statusSheet.getRange("C4").setValue("after coupon OK");
      const img_url = getCreateImg("以下の情報に従って, 未来風にアレンジした画像を生成してください\n\n" + botAnswer);
      statusSheet.getRange("C5").setValue("getCreateImg() OK");

      var requestOptions_text = "";
      if(userLang == "ja"){
        requestOptions_text = "目的地に到着したことを確認しました！\nスタンプを押しました！"  + "\n\n以下の画像は撮影された画像の特徴から未来風にアレンジしたものです！";
      }else if(userLang == "en"){
        requestOptions_text = "We have confirmed that you have arrived at your destination!\nStamped!\nThe following images are arranged in a futuristic style based on the features of the images taken!";
      }

      statusSheet.getRange("C6").setValue("requestOptions_text OK");

      var requestOptions = {
        'headers': {
          'Content-Type': 'application/json; charset=UTF-8',
          'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
        'payload': JSON.stringify({
          'replyToken': replyToken,
          'messages': [{
            'type': 'text',
            'text': requestOptions_text,
          },{
            'type': 'image',
            "originalContentUrl": img_url,
            "previewImageUrl": img_url,
          },{
            'type': 'text',
            'text': matchedMessage,
          }]
        })
      };

      statusSheet.getRange("C7").setValue("set requestOptions OK");

      UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する

      statusSheet.getRange("C8").setValue("fetch OK");

      //スタンプ画像保存
      userStatusSheet.getRange(`${stampCount_column}${userRow}`).setValue(stampCount);
      if(stampCount != 0) userStatusSheet.getRange(`${alp[stampCount_n+stampCount]}${userRow}`).setValue(img_url);

      return returnSuccess();
    }
    else{
      var requestOptions_text;
      if(userLang == "ja"){
        requestOptions_text = "目的地に到着したことを確認できませんでした...\nもう一度撮ってみるじゃん！";
      }else if(userLang == "en"){
        requestOptions_text = "We could not confirm that we had reached our destination...\nPlease take the picture again!";
      }
      var requestOptions = {
        'headers': {
          'Content-Type': 'application/json; charset=UTF-8',
          'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
        'payload': JSON.stringify({
          'replyToken': replyToken,
          'messages': [{
            'type': 'text',
            'text': requestOptions_text,
          }]
        })
      };

      UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
      return returnSuccess();

    }
  }

  statusSheet.getRange("H4").setValue("before system_keywords");

  const system_keywords = [ //定型文の設定
    "さがす", "Search", "さがす: 選択肢", "Search: Choose category", "さがす: AIと相談", "Search: Consult with AI",
    "おわる", "やめとく", "Pass", "スタンプラリー", "Stamp Rally",
    "いく", "english", "日本語"
  ];
  for(var i of system_keywords){ //定型文の判定
    if(userMessage == "さがす: AIと相談"){
      //statusSheet.getRange("A2").setValue("true");//chatGPTとの対話を可能にする
      userStatusSheet.getRange(`${aiAvailablekey_column}${userRow}`).setValue("true");
      const json_button = { //testcode
        'headers': {
            'Content-Type': 'application/json; charset=UTF-8',
            'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
        'payload': JSON.stringify({
            'replyToken': replyToken,
            'messages': [{
              'type': 'template',
              'altText': 'button',
              'template': {
                "type": "buttons",
                "text": "会話を終了するときはこのボタンを押してください！",
                "actions": [
                  {
                    "type": "message",
                    "label": "おわる",
                    "text": "おわる"
                  }//,
                  /*{
                    "type": "postback",
                    "label": "ポストバック",
                    "data":  "test1",
                    'displayText': 'ポストバックが実行されました',
                  }*/
                ]
              }
            }]
        })
      }
      UrlFetchApp.fetch(url, json_button);
      return returnSuccess();
    }
    if(userMessage == "Search: Consult with AI"){
      //statusSheet.getRange("A2").setValue("true");//chatGPTとの対話を可能にする
      userStatusSheet.getRange(`${aiAvailablekey_column}${userRow}`).setValue("true");
      const json_button = { //testcode
        'headers': {
            'Content-Type': 'application/json; charset=UTF-8',
            'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
        'payload': JSON.stringify({
            'replyToken': replyToken,
            'messages': [{
              'type': 'template',
              'altText': 'button',
              'template': {
                "type": "buttons",
                "text": "Press this button to end the conversation!",
                "actions": [
                  {
                    "type": "message",
                    "label": "Finish",
                    "text": "Finish"
                  }
                ]
              }
            }]
        })
      }
      UrlFetchApp.fetch(url, json_button);
      return returnSuccess();
    }
    else if(userMessage == "おわる"){
      userStatusSheet.getRange(`${aiAvailablekey_column}${userRow}`).setValue("false"); //AI相談モードを終了
      userStatusSheet.getRange(`${logCount_column}${userRow}`).setValue(0); //log_countを0
      //会話履歴を削除
      //3+2*log_count -> 3... alp[3]="D", 2*log_count= 一回会話するごとにuserの入力とGPT側の返信が記録されるため,2列ずつ記録が増える
      if(log_count != 0) userStatusSheet.getRange(`${plomptStart_column}${userRow}:${alp[plomptStart_n-1+2*log_count]}${userRow}`).clearContent(); 
      return returnSuccess();
    }
    else if(userMessage == "Finish"){
      userStatusSheet.getRange(`${aiAvailablekey_column}${userRow}`).setValue("false"); //AI相談モードを終了
      userStatusSheet.getRange(`${logCount_column}${userRow}`).setValue(0); //log_countを0
      //会話履歴を削除
      //3+2*log_count -> 3... alp[3]="D", 2*log_count= 一回会話するごとにuserの入力とGPT側の返信が記録されるため,2列ずつ記録が増える
      if(log_count != 0) userStatusSheet.getRange(`${plomptStart_column}${userRow}:${alp[plomptStart_n-1+2*log_count]}${userRow}`).clearContent(); 
      return returnSuccess();
    }
    else if(userMessage.includes("いってみる: ")){
      var destination = userMessage.slice(7); //目的地設定, いってみる: で6文字,それ以降を切り取るため,7文字目から
      userStatusSheet.getRange(`${destination_column}${userRow}`).setValue(destination);
      var requestOptions = setRequestOptions(`目的地を「${destination}」に設定しました！\n画面下のメニューの中央にある「いく」をタップすると経路を案内します！\nまた, メニュー右側にある「とる」をタップするとカメラを起動します！写真を送信してポイントゲット！`, replyToken);
      UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
      return returnSuccess();
    }
    else if(userMessage.includes("Go: ")){
      var destination = userMessage.slice(3); //目的地設定, Go: で3文字,それ以降を切り取るため,4文字目から
      userStatusSheet.getRange(`${destination_column}${userRow}`).setValue(destination);
      var requestOptions = setRequestOptions(`The destination is now set to "${destination}"! \nTap Go in the center of the menu at the bottom of the screen to guide you along the route! \nAlso, tap Take on the right side of the menu to start the camera! Send photos and get points!`, replyToken);
      UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
      return returnSuccess();
    }
    else if(userMessage == "いく"){
      const destination = userStatusSheet.getRange(`${destination_column}${userRow}`).getValue(); //目的地を取得
      statusSheet.getRange("I1").setValue("get destination OK");
      var requestOptions;
      if(destination == ""){
        requestOptions = setRequestOptions(`目的地が未設定のようです！\n目的地は「さがす」から設定できます！` ,replyToken);
      }else{
        const map_url = getMapUrl("甲府駅", destination, "public"); 
        requestOptions = setRequestOptions(`目的地「${destination}」への経路は以下の通りです！\n${map_url}` ,replyToken);
      }
      UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
      return returnSuccess();
    }
    else if(userMessage == "Go"){
      const destination = userStatusSheet.getRange(`${destination_column}${userRow}`).getValue(); //目的地を取得
      statusSheet.getRange("I1").setValue("get destination OK");
      var requestOptions;
      if(destination == ""){
        requestOptions = setRequestOptions(`The destination seems to be unset! \nYou can set your destination from "Search"!` ,replyToken);
      }else{
        const map_url = getMapUrl("甲府駅", destination, "public"); 
        requestOptions = setRequestOptions(`The route to the destination "${destination}" is as follows \n${map_url}` ,replyToken);
      }
      UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
      return returnSuccess();
    }
    else if(userMessage == "スタンプラリー" || userMessage == "Stamp Rally"){
      const stampCount = Number(userStatusSheet.getRange(`${stampCount_column}${userRow}`).getValue());
      var requestOptions = {
        'headers': {
            'Content-Type': 'application/json; charset=UTF-8',
            'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
        'payload': JSON.stringify({
          'replyToken': replyToken,
          'messages': [{
            "type": "template",
            "altText": "this is a image carousel template",
            "template": {
              "type": "image_carousel",
              "columns": setColumns(stampCount, userRow),
            }
          }]
        })
      }
    UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
    return returnSuccess();
    }
    else if(userMessage == "english"){
      userStatusSheet.getRange(`${lang_column}${userRow}`).setValue("en");
      var requestOptions = {
        'headers': {
          'Content-Type': 'application/json; charset=UTF-8',
          'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
      }
      UrlFetchApp.fetch(`https://api.line.me/v2/bot/user/${userID}/richmenu/${richMenuId_en}`, requestOptions);
      return returnSuccess();
    }
    else if(userMessage == "日本語"){
      userStatusSheet.getRange(`${lang_column}${userRow}`).setValue("ja");
      var requestOptions = {
        'headers': {
          'Content-Type': 'application/json; charset=UTF-8',
          'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN,
        },
        'method': 'post',
      }
      UrlFetchApp.fetch(`https://api.line.me/v2/bot/user/${userID}/richmenu/${richMenuId_ja}`, requestOptions);
      return returnSuccess();
    }
    else if(userMessage.includes("やめとく: ") || userMessage.includes("Pass: ") || userMessage.includes("カテゴリ: ") || userMessage.includes("Category: ")){
      return returnSuccess();
    }
    if(userMessage == i){ //上記以外の定型文の場合
      return returnSuccess();
    }
  }

  if(userStatusSheet.getRange(`${aiAvailablekey_column}${userRow}`).getValue() == false){ //AI相談モードでなければ終了
    var requestOptions_text;
    if(userLang == "ja"){
      requestOptions_text = 'メッセージありがとうございます！\nサービスの利用は画面下のメニューからお願いします！';
    }else if(userLang == "en"){
      requestOptions_text = 'Thank you for your message!\nPlease use the menu at the bottom of the screen to access our services!';
    }
    requestOptions = setRequestOptions(requestOptions_text, replyToken);
    UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
    return returnSuccess();
  } 

  //chatGPTとの対話の履歴の取得
  userStatusSheet.getRange(`${logCount_column}${userRow}`).setValue(log_count + 1);
  let prompt_assistant = [];  //chatGPT側(assistant)の返信群
  let prompt_user = []; //user側の入力群
  for(var i=0; i<log_count; i++){ //実際に履歴を取得(奇数行はuser側, 偶数行はchatGPT側)
    statusSheet.getRange("B4").setValue("for in OK");
    prompt_user.push(userStatusSheet.getRange(`${alp[plomptStart_n+2*i]}${userRow}`).getValue());
    prompt_assistant.push(userStatusSheet.getRange(`${alp[plomptStart_n+2*i+1]}${userRow}`).getValue());
  }
  prompt_user.push(userMessage); //最後にユーザーの入力を配列に保存
  statusSheet.getRange("B5").setValue("for OK");
  const gptReply = getReply(log_count, prompt_user, prompt_assistant); //chatGPTと通信
  statusSheet.getRange("B6").setValue("chatGPT OK");
  const botMessage = gptReply.message; //chatGPTの返信メッセージ
  statusSheet.getRange("B8").setValue("bot message save OK");
  appendLog_status(userRow, log_count, userMessage, botMessage); //対話履歴を保存
  statusSheet.getRange("B10").setValue("appdendLog_status OK");


  var requestOptions = setRequestOptions(botMessage, replyToken); //送信データの用意 
  const response = UrlFetchApp.fetch(url, requestOptions); //LINE側にデータを送信する
  
  //appendLog関連
  const timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const userId = event.source.userId;
  const messageType = event.message.type;
  let errorMessage = "";
  if(response !== {}){
    errorMessage = response.message;
  }
  const promptTokens = gptReply.usage.prompt_tokens;
  const completionTokens = gptReply.usage.completion_tokens;
  const totalTokens = gptReply.usage.total_tokens;

  appendLog([
    timestamp,
    userId,
    messageType,
    userMessage,
    botMessage,
    promptTokens,
    completionTokens,
    totalTokens,
    errorMessage
  ]);

  return returnSuccess();
}