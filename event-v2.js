/*
 * EVENT-V2 NEUTRALNY
 * WERSJA: EVENT-V2 NEUTRAL 2026-06-01 01
 * Ten plik nie jest wskazywany przez aktualny manifest. Zostaje neutralny na wypadek starego cache.
 */

function onNewMessageComposeHandlerV2(event) {
  try {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  } catch (e) {
  }
}

try {
  Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
} catch (e) {
}
