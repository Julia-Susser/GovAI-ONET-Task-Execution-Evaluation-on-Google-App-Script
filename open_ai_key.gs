

function callOpenAI(promptText) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: 'gpt-3.5-turbo',
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
  
  const requests = prompts.map(prompt => {
    var payload = {
    model: 'gpt-3.5-turbo',
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
  console.log(requests)
  try {
    const responses = UrlFetchApp.fetchAll(requests);
    texts = responses.map(response => {
      var json = JSON.parse(response.getContentText());
      var text = json.choices[0].message.content.trim();
      return [text]
    })
    return texts
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error;
  }
}