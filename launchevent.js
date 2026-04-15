/* global Office */

function onMessageSendHandler(event) {
  // TODO: ここにあなたの宛先ドメイン判定ロジックを入れる
  // NGなら event.completed({ allowEvent: false, errorMessage: "...", ... })
  // OKなら event.completed({ allowEvent: true })

  event.completed({ allowEvent: true });
}

function action(event) {
  event.completed();
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("action", action);