function onNewMessageComposeHandler(event) {
    try {
        event.completed();
    } catch (e) {
    }
}

Office.actions.associate(
    "onNewMessageComposeHandler",
    onNewMessageComposeHandler
);
