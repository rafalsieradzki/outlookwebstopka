function onNewMessageComposeHandlerV2(event) {
    try {
        event.completed();
    } catch (e) {
    }
}

Office.actions.associate(
    "onNewMessageComposeHandlerV2",
    onNewMessageComposeHandlerV2
);
