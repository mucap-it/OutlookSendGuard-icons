/* global Office */

/**
 * ===== 設定 =====
 * 社内扱いのドメイン（複数可）
 */
var INTERNAL_DOMAINS = [
  "mucap.co.jp" // TODO: 自社ドメインに変更/追加
];

var MAX_LIST = 8;

/** ===== util ===== */
function normalizeEmail(email) {
  return (email || "").trim().toLowerCase();
}

function getDomain(email) {
  var e = normalizeEmail(email);
  var at = e.lastIndexOf("@");
  return at >= 0 ? e.substring(at + 1) : "";
}

function isExternal(email) {
  var domain = getDomain(email);
  if (!domain) return false;
  return INTERNAL_DOMAINS.indexOf(domain) === -1;
}

function uniq(arr) {
  var map = {};
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var v = arr[i];
    if (!map[v]) {
      map[v] = true;
      out.push(v);
    }
  }
  return out;
}

function getRecipientsAsync(field, callback) {
  if (!field || typeof field.getAsync !== "function") {
    callback([]);
    return;
  }
  field.getAsync(function (result) {
    callback((result && result.value) ? result.value : []);
  });
}

/**
 * ===== Smart Alerts: OnMessageSend =====
 * - 外部宛が1件でもあれば allowEvent:false で Smart Alerts ダイアログを出す（SoftBlock）
 */
function onMessageSendHandler(event) {
  try {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) {
      // 取得できない場合は業務影響回避で送信許可
      event.completed({ allowEvent: true });
      return;
    }

    getRecipientsAsync(item.to, function (toList) {
      getRecipientsAsync(item.cc, function (ccList) {
        getRecipientsAsync(item.bcc, function (bccList) {

          var all = []
            .concat(toList || [])
            .concat(ccList || [])
            .concat(bccList || []);

          var external = uniq(
            all
              .map(function (r) { return normalizeEmail(r.emailAddress); })
              .filter(function (addr) { return addr && isExternal(addr); })
          );

          if (external.length > 0) {
            var listed = external.slice(0, MAX_LIST);
            var remaining = external.length - listed.length;

            var plain =
              "社外宛のメールアドレスが含まれています。宛先に誤りがないかご確認ください。\n\n" +
              listed.join("\n") +
              (remaining > 0 ? ("\n…ほか " + remaining + " 件") : "");

            // SoftBlock: allowEvent:false でダイアログ表示（送信続行はユーザー判断）
            event.completed({
              allowEvent: false,
              errorMessage: plain,
              cancelLabel: "編集に戻る"
            });
            return;
          }

          // 問題なし → 送信許可
          event.completed({ allowEvent: true });
        });
      });
    });

  } catch (e) {
    // 例外時も業務影響回避で送信許可
    event.completed({ allowEvent: true });
  }
}

/**
 * 既存のリボンボタン(action)用（残しておく：デバッグに便利）
 */
function action(event) {
  try {
    var msg = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Performed action.",
      icon: "Icon.80x80",
      persistent: true
    };
    if (Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", msg);
    }
  } catch (e) { /* noop */ }

  event.completed();
}

/**
 * ★必須：manifest.xml の FunctionName と一致させる
 * - LaunchEvent の FunctionName="onMessageSendHandler"
 * - ボタンの FunctionName="action"
 */
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("action", action);