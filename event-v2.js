// Nieużywany plik po uproszczeniu do jednego eventu.
// Zostawiony neutralnie, aby stare cache nie wstawiały żadnych testowych treści.

function onNewMessageComposeHandlerV2(event) {
  try { event.completed(); } catch (e) {}
}

Office.actions.associate("onNewMessageComposeHandlerV2", onNewMessageComposeHandlerV2);
