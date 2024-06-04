

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