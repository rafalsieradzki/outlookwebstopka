/*
 * EVENT-V2 NEUTRALNY
 * WERSJA: EVENT-V2 NEUTRAL 2026-06-01 02
 * Ten plik nic nie wstawia. Zostaje neutralny na wypadek starego cache Outlooka.
 */

function completeEventV2(event) {
  try {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (e) {
  }
}

function onNewMessageComposeHandlerV2(event) {
  completeEventV2(event);
}

try {
  Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
} catch (e) {
}
