const NOTIFY_TO = "ntst118@hbg.ac.jp";
const EMAIL_SUBJECT_PREFIX = "フォーム回答通知";
const PROPERTY_FORM_TITLE = "FORM_TITLE";
const PROPERTY_SHEET_URL = "SHEET_URL";

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

  const formInfo = getFormInfoFromEvent_(e);
  const formTitle = formInfo.formTitle;
  const sheetUrl = formInfo.sheetUrl;
  const subject = `${EMAIL_SUBJECT_PREFIX}: ${formTitle}`;
  const body = buildMailBody_(e, formTitle, sheetUrl);

  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject,
    body,
  });
}

function setupOnFormSubmitTrigger() {
  const form = FormApp.getActiveForm();
  if (!form) {
    return;
  }

  resetOnFormSubmitTrigger_(form);
  storeFormInfo_(form);
}

function setupOnFormSubmitTriggerFromSheet() {
  const form = getLinkedFormFromSheet_();
  if (!form) {
    SpreadsheetApp.getUi().alert("このスプレッドシートに紐づくフォームが見つかりませんでした。");
    return;
  }

  resetOnFormSubmitTrigger_(form);
  storeFormInfo_(form);

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

function storeFormInfo_(form) {
  const properties = PropertiesService.getScriptProperties();
  const sheetUrl = getLinkedSpreadsheetUrl_(form);
  const formTitle = normalizeTitle_(form.getTitle());

  if (formTitle) {
    properties.setProperty(PROPERTY_FORM_TITLE, formTitle);
  }
  if (sheetUrl) {
    properties.setProperty(PROPERTY_SHEET_URL, sheetUrl);
  }
}

function getStoredFormInfo_() {
  const properties = PropertiesService.getScriptProperties();

  return {
    formTitle: properties.getProperty(PROPERTY_FORM_TITLE) || "",
    sheetUrl: properties.getProperty(PROPERTY_SHEET_URL) || "",
  };
}

function buildMailBody_(e, formTitle, sheetUrl) {
  const lines = [];
  const response = e.response;
  const namedValues = e.namedValues;
  const safeFormTitle = normalizeTitle_(formTitle) || "フォーム";

  lines.push(`フォーム名: ${safeFormTitle}`);
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
  if (sheetUrl) {
    lines.push("");
    lines.push(`スプレッドシートを開く: ${sheetUrl}`);
  }

  return lines.join("\n");
}

function getFormInfoFromEvent_(e) {
  const source = e.source;
  let formTitle = "";
  let sheetUrl = "";

  if (source) {
    if (typeof source.getDestinationType === "function") {
      // フォームの送信トリガー
      formTitle = normalizeTitle_(source.getTitle());
      sheetUrl = getLinkedSpreadsheetUrl_(source);
    } else if (typeof source.getFormUrl === "function") {
      // スプレッドシート側のフォーム送信トリガー
      sheetUrl = source.getUrl();
      const formUrl = source.getFormUrl();
      if (formUrl) {
        try {
          const form = FormApp.openByUrl(formUrl);
          formTitle = normalizeTitle_(form.getTitle());
        } catch (error) {
          if (source.getTitle) {
            formTitle = normalizeTitle_(source.getTitle()) || formTitle;
          }
        }
      } else if (source.getTitle) {
        formTitle = normalizeTitle_(source.getTitle()) || formTitle;
      }
    } else if (source.getTitle) {
      formTitle = normalizeTitle_(source.getTitle()) || formTitle;
    }
  }

  const storedInfo = getStoredFormInfo_();
  if (!formTitle && storedInfo.formTitle) {
    formTitle = storedInfo.formTitle;
  }
  if (!sheetUrl && storedInfo.sheetUrl) {
    sheetUrl = storedInfo.sheetUrl;
  }

  return { formTitle: formTitle || "フォーム", sheetUrl };
}

function normalizeTitle_(title) {
  if (typeof title !== "string") {
    return "";
  }
  const trimmed = title.trim();
  return trimmed ? trimmed : "";
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

function resetOnFormSubmitTrigger_(form) {
  // 既存トリガーを整理してから再作成する
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === "onFormSubmit")
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();
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
