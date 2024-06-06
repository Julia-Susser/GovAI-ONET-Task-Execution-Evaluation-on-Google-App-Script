

function callOpenAI(promptText) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: 'gpt-4',
    messages: [{"role": "user", "content": promptText}],
    max_tokens: 100
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const text = json.choices[0].message.content.trim();
    return text;
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error;
  }
}


function batchOpenAIRequest(prompts){
  const url = 'https://api.openai.com/v1/chat/completions';
  var chunkSize = 10
  const requests = prompts.map(prompt => {
    var payload = {
    model: 'gpt-4',
    messages: [{"role": "user", "content": prompt}],
    max_tokens: 100
    };
  
    return {
    url : url,
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload)
  }} )
  try {
    var texts = []
    for (let i = 0; i < requests.length; i += chunkSize) {
        const chunk = requests.slice(i, i + chunkSize);
        const responses = UrlFetchApp.fetchAll(chunk);
        batchTexts = responses.map(response => {
          var json = JSON.parse(response.getContentText());
          var text = json.choices[0].message.content.trim();
          return [text]
        })
        texts = texts.concat(batchTexts);
    }
    return texts
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error;
  }
}