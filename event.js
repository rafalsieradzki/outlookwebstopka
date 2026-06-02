function onNewMessageComposeHandler(event){try{event.completed();}catch(e){}}
try{Office.actions.associate("onNewMessageComposeHandler",onNewMessageComposeHandler);}catch(e){}
