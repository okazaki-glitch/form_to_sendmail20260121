const NOTIFY_TO = "ntst118@hbg.ac.jp";
const EMAIL_SUBJECT_PREFIX = "フォーム回答通知";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("フォーム通知")
    .addItem("トリガー再設定", "setupOnFormSubmitTriggerFromSheet")
    .addToUi();
}

function onFormSubmit(e) {
  if (!e) {
    return;
  }

  const form = e.source;
  const formTitle = form ? form.getTitle() : "フォーム";
  const sheetUrl = getLinkedSpreadsheetUrl_(form);
  const subject = `${EMAIL_SUBJECT_PREFIX}: ${formTitle}`;
  const body = buildMailBody_(e, formTitle, sheetUrl);

  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject,
    body,
  });
}

function setupOnFormSubmitTrigger() {
  // 既存トリガーを整理してから再作成する
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === "onFormSubmit")
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(FormApp.getActiveForm())
    .onFormSubmit()
    .create();
}

function setupOnFormSubmitTriggerFromSheet() {
  const form = getLinkedFormFromSheet_();
  if (!form) {
    SpreadsheetApp.getUi().alert("このスプレッドシートに紐づくフォームが見つかりませんでした。");
    return;
  }

  // 既存トリガーを整理してから再作成する
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === "onFormSubmit")
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert("フォーム送信トリガーを再設定しました。");
}

function getLinkedFormFromSheet_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const formUrl = sheet.getFormUrl();
    if (!formUrl) {
      return null;
    }
    return FormApp.openByUrl(formUrl);
  } catch (error) {
    return null;
  }
}

function buildMailBody_(e, formTitle, sheetUrl) {
  const lines = [];
  const response = e.response;
  const namedValues = e.namedValues;

  lines.push(`フォーム名: ${formTitle}`);
  if (sheetUrl) {
    lines.push(`スプレッドシートを開く: ${sheetUrl}`);
  }
  if (response) {
    const timestamp = response.getTimestamp();
    if (timestamp) {
      lines.push(`送信日時: ${timestamp}`);
    }
    const respondentEmail = response.getRespondentEmail();
    if (respondentEmail) {
      lines.push(`回答者メール: ${respondentEmail}`);
    }
  }

  lines.push("");
  lines.push("回答内容:");
  lines.push(formatNamedValues_(namedValues, response));

  return lines.join("\n");
}

function getLinkedSpreadsheetUrl_(form) {
  if (!form) {
    return "";
  }

  try {
    if (form.getDestinationType() !== FormApp.DestinationType.SPREADSHEET) {
      return "";
    }
    const spreadsheetId = form.getDestinationId();
    if (!spreadsheetId) {
      return "";
    }
    return SpreadsheetApp.openById(spreadsheetId).getUrl();
  } catch (error) {
    return "";
  }
}

function formatNamedValues_(namedValues, response) {
  if (namedValues && Object.keys(namedValues).length > 0) {
    return Object.keys(namedValues)
      .map((question) => {
        const answers = namedValues[question];
        const answerText = Array.isArray(answers) ? answers.join(", ") : answers;
        return `- ${question}: ${answerText}`;
      })
      .join("\n");
  }

  if (response) {
    const itemResponses = response.getItemResponses();
    if (itemResponses.length > 0) {
      return itemResponses
        .map((itemResponse) => {
          const question = itemResponse.getItem().getTitle();
          const answer = itemResponse.getResponse();
          const answerText = Array.isArray(answer) ? answer.join(", ") : answer;
          return `- ${question}: ${answerText}`;
        })
        .join("\n");
    }
  }

  return "(回答データが取得できませんでした)";
}
