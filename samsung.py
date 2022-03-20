import requests
 
def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )
    print(response)
 
myToken = "xoxb-3031085882624-3267050540002-QTFXuyz9RM4AlBITRrQPosDf"
 
post_message(myToken,"#stock-info","testetest")