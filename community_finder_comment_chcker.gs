const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let url, fabicon_url, webhook_url, mentions;

function getCredentials() {
  let credentials_sheet = spreadsheet.getSheetByName('Credentials');
  let mentions_sheet = spreadsheet.getSheetByName('UserIDs');
  url = credentials_sheet.getRange('A2').getValue();
  webhook_url = credentials_sheet.getRange('B2').getValue();
  fabicon_url = credentials_sheet.getRange('C2').getValue();

  let collength = mentions_sheet.getRange(1,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getColumn();
  mentions = mentions_sheet.getRange(1, 1, 1, collength).getValues();
}

function main() {
  getCredentials();
  let sheets = spreadsheet.getSheets();
  let latest_sheet_date = sheets.splice(3).map(s => new Date(s.getName())).sort(function(a,b) { return (a > b ? 1 : -1) }).pop();

  let sheet = putCommunityFinderComments();
  let comments = sheet.getDataRange().getValues().splice(1);

  let new_comments = comments.filter(c => c[2] > latest_sheet_date);

  let initialize_message = {
    "content": "",
    "embeds": [
      {
        "title": "Freesia CommunityFinder Comment Checker",
        "fields": [
          {
            "name": ":stopwatch: Triggered DateTime",
            "value": Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
            "inline": true,
          },
        ],
      },
    ],
  };
  sendDiscordMessage(initialize_message);

  let message;
  if (new_comments.length > 0) {
    let embeds = new_comments.map(c => {
      return {
        "title": "新着メッセージ",
        "url": url,
        "thumbnail": {"url": c[3]},
        "color": parseInt("faa61a", 16),
        "fields": [
          {
            "name": ":memo: プレイヤー名",
            "value": c[0],
            "inline": false
          },
          {
            "name": ":mega: コメント",
            "value": c[1],
            "inline": false,
          },
          {
            "name": ":clock1: 投稿時間",
            "value": Utilities.formatDate(c[2], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
            "inline": false,
          },
        ],
      }
    });

    let content_text = mentions.join('\n') + "\n新着メッセージがあります！";
    message = {
      "content": content_text,
      "tts": false,
      "embeds": embeds,
    };
  } else {
    message = {
      "tts": false,
      "embeds": [
        {
          "description": "新着メッセージはありませんでした",
          "color": parseInt("43b581", 16),
        }
      ]
    }
  }

  sendDiscordMessage(message);
}

function putCommunityFinderComments() {
  let response = UrlFetchApp.fetch(url);
  let text = response.getContentText('utf-8');

  const $ = Cheerio.load(text);
  let sheet = createNewSheet();

  $('[class="cf-comment"]').each((_, elem) => {
    let comment_parent_body = $(elem).children('.cf-comment__text.js__comment');
    let comment_body = $(comment_parent_body).children('.cf-comment__bg');
    let name_body = $(comment_body).children('.cf-comment__header').children('.cf-comment__name');

    if (! $(name_body).children('.member.cf-color').length > 0) {
      let face_url = $(elem).children('.cf-comment__face').children('a').children('img').attr('src');
      let name = getSurfaceText(name_body).replace(/\r?\n|\s+/g, '');
      let comment = $(comment_body).children('.cf-comment__body').html().replace(/\s+/g, '').replace(/<br>/g, '\n');
      let unix_time = $(comment_parent_body).children('.cf-comment__data').children('.datetime_dynamic_ymdhm').attr('data-epoch');
      let created_at = Utilities.formatDate(new Date(unix_time * 1000), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

      sheet.appendRow([name, comment, created_at, face_url]);
    }
  });

  return sheet;
}

function getSurfaceText(cheerio){
    cheerio.children().empty();
    return cheerio.text();
}

function createNewSheet() {
  let date = new Date();
  let template = spreadsheet.getSheetByName('Template');
  let sheet = template.copyTo(spreadsheet);
  sheet.setName(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));

  return sheet;
}

function sendDiscordMessage(payload) {
    let param = {
    "method": "POST",
    "headers": { "Content-type": "application/json" },
    "payload": JSON.stringify(payload),
  };

  try {
    UrlFetchApp.fetch(webhook_url, param);
  } catch (e) {
    let message = {
      "tts": false,
      "embeds": [
        {
          "title": "ERROR",
          "description": "トリガー実行中にエラーが発生しました。エラーログを確認してください。",
          "color": parseInt("f04747", 16),
        }
      ]
    };
    let param = {
      "method": "POST",
      "headers": { "Content-type": "application/json" },
      "payload": JSON.stringify(message),
    };
    UrlFetchApp.fetch(webhook_url, param);
    throw e;
  }
}
