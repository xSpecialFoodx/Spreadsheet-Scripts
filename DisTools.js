// user config start

// the webhook link of the dis channel
var DisWebHook = "https://dis.com/api/webhooks/numbers/text";

// user config end

var DisToolsObject = null;// do not delete

// whether to send dis messages if running in dry run
var send_dis_messages_in_dry_run = false;// do not delete

// logging the responses of the calls to the dis api
var LogDisMessagesResponses = false;// do not delete

// the dis text styles start

var ItalicsDisTextStyle = (
  {
    Code: 1//parseInt("0001", 2)
    , Symbol: "*"
  }
);// do not delete

var BoldDisTextStyle = (
  {
    Code: 2//parseInt("0010", 2)
    , Symbol: "**"
  }
);// do not delete

var UnderlineDisTextStyle = (
  {
    Code: 4//parseInt("0100", 2)
    , Symbol: "__"
  }
);// do not delete

var StrikethroughDisTextStyle = (
  {
    Code: 8//parseInt("1000", 2)
    , Symbol: "~~"
  }
);// do not delete

// the dis text styles end

class DisToolsObjectType
{
  constructor()
  {
    InitializeGeneralToolsObject_();
  }
  
  // sending a dis message
  SendDisMessage(Message)
  {
    if (
      Message != null
      && GeneralToolsObject.VariableIsString(Message) == true
      && Message.length > 0
    )
      if (
        DisWebHook != undefined
        && DisWebHook != null
        && GeneralToolsObject.VariableIsString(DisWebHook) == true
        && DisWebHook.length > 0
      )
        if (dry_run == false || send_dis_messages_in_dry_run == true)
        {
          var Response =
            UrlFetchApp.fetch(
              DisWebHook
              , (
                {
                  method: "POST",
                  payload: JSON.stringify({content: Message}),
                  muteHttpExceptions: true,
                  contentType: "application/json"
                }
              )
            );

          if (LogDisMessagesResponses == true)
            Logger.log(Response.getContentText());
        }
  }

  // applying a text style to a dis text
  ApplyTextStyle(Text, TextStyleCode)
  {
    var FunctionResult;
    var CurrentResult = null;

    if (
      Text != null
      && GeneralToolsObject.VariableIsString(Text) == true
      && Text.length > 0
    )
    {
      CurrentResult = Text;

      if (
        TextStyleCode != null
        && GeneralToolsObject.VariableIsNumber(TextStyleCode) == true
        && GeneralToolsObject.VariableIsString(TextStyleCode) == false
        && TextStyleCode > 0
        && TextStyleCode <= (
          ItalicsDisTextStyle.Code
          | BoldDisTextStyle.Code
          | UnderlineDisTextStyle.Code
          | StrikethroughDisTextStyle.Code
        )
      )
      {
        if ((TextStyleCode & ItalicsDisTextStyle.Code) > 0)
          CurrentResult = ItalicsDisTextStyle.Symbol + CurrentResult + ItalicsDisTextStyle.Symbol;

        if ((TextStyleCode & BoldDisTextStyle.Code) > 0)
          CurrentResult = BoldDisTextStyle.Symbol + CurrentResult + BoldDisTextStyle.Symbol;

        if ((TextStyleCode & UnderlineDisTextStyle.Code) > 0)
          CurrentResult = UnderlineDisTextStyle.Symbol + CurrentResult + UnderlineDisTextStyle.Symbol;

        if ((TextStyleCode & StrikethroughDisTextStyle.Code) > 0)
          CurrentResult = StrikethroughDisTextStyle.Symbol + CurrentResult + StrikethroughDisTextStyle.Symbol;
      }
    }

    FunctionResult = CurrentResult;

    return FunctionResult;
  }

  // check the format to a dis member ping
  CheckPingDisMemberFormat(UserID)
  {
    return (
      UserID != null
      && GeneralToolsObject.VariableIsString(UserID) == true
      && UserID.length > 0
      ? '<' + '@' + UserID + '>'
      : null
    );
  }
}

// initializing the dis tools object
function InitializeDisToolsObject_()
{
  if (DisToolsObject == null)
    DisToolsObject = new DisToolsObjectType();
}
