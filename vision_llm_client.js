const fs = require('fs');
const path = require('path');

/**
 * 讀取設定檔並呼叫 Vision LLM
 * 支援: openai, gemini, ollama
 */
async function callVisionLLM(prompt, imagePaths) {
  const configPath = path.join(__dirname, '.vision_llm_config.json');
  if (!fs.existsSync(configPath)) {
    throw new Error(`Config file not found at ${configPath}`);
  }

  const config = JSON.parse(fs.readFileSync(configPath, 'utf8').replace(/^\uFEFF/, ''));
  const provider = config.provider || 'openai';
  const providerConfig = config[provider];

  if (!providerConfig) {
    throw new Error(`Provider configuration for "${provider}" not found.`);
  }

  // 取得 API Key
  let apiKey = providerConfig.api_key;
  if (!apiKey && providerConfig.api_key_env) {
    apiKey = process.env[providerConfig.api_key_env];
  }

  // 將圖片轉為 Base64
  const imageMessages = imagePaths.map(p => {
    const data = fs.readFileSync(p).toString('base64');
    const mime = p.endsWith('.png') ? 'image/png' : 'image/jpeg';
    return { data, mime };
  });

  if (provider === 'openai') {
    return await callOpenAI(providerConfig, apiKey, prompt, imageMessages);
  } else if (provider === 'fireworks') {
    return await callOpenAICompatible(providerConfig, apiKey, prompt, imageMessages, 'https://api.fireworks.ai/inference/v1/chat/completions');
  } else if (provider === 'gemini') {
    return await callGemini(providerConfig, apiKey, prompt, imageMessages);
  } else if (provider === 'ollama') {
    return await callOllama(providerConfig, apiKey, prompt, imageMessages);
  } else if (provider === 'custom') {
    return await callOpenAICompatible(providerConfig, apiKey, prompt, imageMessages);
  } else {
    throw new Error(`Unsupported provider: ${provider}`);
  }
}

async function callOpenAI(config, apiKey, prompt, images) {
  const content = [{ type: 'text', text: prompt }];
  images.forEach(img => {
    content.push({
      type: 'image_url',
      image_url: { url: `data:${img.mime};base64,${img.data}` }
    });
  });

  const response = await postJson('https://api.openai.com/v1/chat/completions', {
    model: config.model || 'gpt-4o',
    messages: [{ role: 'user', content }],
    max_tokens: config.max_tokens || 4096
  }, {
    Authorization: `Bearer ${apiKey}`,
    'Content-Type': 'application/json'
  });

  return response.choices[0].message.content;
}

async function callOpenAICompatible(config, apiKey, prompt, images, defaultUrl) {
  const endpoint = config.endpoint_url || defaultUrl;
  if (!endpoint) {
    throw new Error('OpenAI-compatible provider requires endpoint_url.');
  }

  const content = [{ type: 'text', text: prompt }];
  images.forEach(img => {
    content.push({
      type: 'image_url',
      image_url: { url: `data:${img.mime};base64,${img.data}` }
    });
  });

  const body = {
    model: config.model,
    messages: [
      {
        role: 'system',
        content: config.system_prompt || 'You are a strict JSON API. Return only one valid JSON object. Do not include analysis, Markdown, or prose.'
      },
      { role: 'user', content }
    ],
    max_tokens: config.max_tokens || 4096,
    temperature: config.temperature ?? 0
  };

  if (config.response_format) {
    body.response_format = config.response_format;
  }

  const headers = { 'Content-Type': 'application/json' };
  if (apiKey) {
    headers.Authorization = `Bearer ${apiKey}`;
  }

  const response = await postJson(endpoint, body, headers);
  return response.choices[0].message.content;
}

async function callGemini(config, apiKey, prompt, images) {
    // 這裡使用簡單的 REST API 呼叫，不依賴 SDK 以保持輕量
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model}:generateContent?key=${apiKey}`;
    
    const parts = [{ text: prompt }];
    images.forEach(img => {
        parts.push({
            inline_data: { mime_type: img.mime, data: img.data }
        });
    });

    const response = await postJson(url, {
        contents: [{ parts }]
    });

    return response.candidates[0].content.parts[0].text;
}

async function callOllama(config, apiKey, prompt, images) {
  // 許多人使用 local Ollama
  const url = config.endpoint_url || 'http://localhost:11434/api/chat';
  
  const content = prompt;
  const image_data = images.map(img => img.data);

  const response = await postJson(url, {
    model: config.model,
    messages: [{
      role: 'user',
      content: content,
      images: image_data
    }],
    stream: false
  }, apiKey ? { Authorization: `Bearer ${apiKey}` } : {});

  return response.message.content;
}

async function postJson(url, body, headers = {}) {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...headers },
    body: JSON.stringify(body)
  });

  const text = await response.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch (error) {
    throw new Error(`Non-JSON response from ${url}: ${text.slice(0, 500)}`);
  }

  if (!response.ok) {
    const detail = data.error?.message || data.message || text.slice(0, 500);
    throw new Error(`HTTP ${response.status} from ${url}: ${detail}`);
  }

  return data;
}

module.exports = { callVisionLLM };
