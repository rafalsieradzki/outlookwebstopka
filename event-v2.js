function onNewMessageComposeHandlerV2(event){try{event.completed();}catch(e){}}
try{Office.actions.associate("onNewMessageComposeHandlerV2",onNewMessageComposeHandlerV2);}catch(e){}
