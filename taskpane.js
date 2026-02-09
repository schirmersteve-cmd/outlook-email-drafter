// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Email Drafter loaded successfully");
        loadSettings();
    }
});

// Tone prompts for different email styles
const tonePrompts = {
    professional: "Rewrite the following notes as a polished, professional business email. Use formal language, proper structure, and a respectful tone. Keep it clear and concise.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice. Do not add commitments, promises, or technical claims.",
    
    friendly: "Rewrite the following notes as a warm, friendly professional email. Be approachable and personable while maintaining professionalism. Use a conversational but business-appropriate tone.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice.",
    
    casual: "Rewrite the following notes as a casual, relaxed email. Use a conversational tone as if writing to a colleague or familiar contact. Keep it friendly and informal while remaining appropriate for business communication.\n\nPreserve original meaning and factual content. Do not invent details.",
    
    brief: "Rewrite the following notes as a brief, direct email. Get straight to the point. Use short sentences and minimal pleasantries. Be clear and action-oriented.\n\nPreserve original meaning and factual content. Do not invent details.",
    
    diplomatic: "Rewrite the following notes as a diplomatic, tactful email. Handle the situation delicately, acknowledge concerns when appropriate, and maintain a professional and empathetic tone. Focus on solutions and constructive next steps.\n\nPreserve original meaning and factual content. Do not invent details. Do not admit fault or liability unless explicitly stated in the notes.",
    
    cleanup: "Clean up the following notes or draft email.\n\nFix grammar and spelling.\nRemove unnecessary fluff.\nImprove readability and flow.\nNormalize formatting.\nKeep the message concise.\n\nDo NOT significantly change tone or meaning. Preserve original factual content.",
    
    sales: "Rewrite the following notes as a professional, customer-facing business email.\n\nImprove clarity, persuasion, and engagement while maintaining credibility and technical accuracy. Keep tone confident but not pushy.\n\nPreserve original meaning and factual content. Do not invent specifications, performance claims, or commitments.",
    
    myvoice: "Convert the following rough notes, shorthand, or bullet points into a complete professional email written in the sender's natural voice.\n\nGuidelines:\n• Write in a style that is professional but conversational.\n• Be clear, direct, and relationship-focused.\n• Maintain technical credibility without sounding overly formal or marketing-driven.\n• Keep the message natural and easy to read.\n\nContent Rules:\n• Preserve original meaning and factual content.\n• Do not invent details, specifications, or commitments.\n• If notes are vague, keep language appropriately general and safe.\n\nStructure:\n• Organize the message into a logical email flow.\n• Add transitions and readability improvements where needed.\n• Keep length appropriate to the content (do not over-expand)."
};

// Subject line prompts (only used when subject is empty)
const subjectPrompts = {
    professional: "\n\nAlso provide a professional, clear subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    friendly: "\n\nAlso provide a friendly, engaging subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    casual: "\n\nAlso provide a casual, straightforward subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    brief: "\n\nAlso provide a brief, direct subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    diplomatic: "\n\nAlso provide a diplomatic, professional subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    cleanup: "\n\nAlso provide a clear subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    sales: "\n\nAlso provide a compelling, professional subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body.",
    myvoice: "\n\nAlso provide an appropriate subject line for this email. Return it on the first line as 'Subject: [your subject]' followed by a blank line, then the email body."
};

// View Management
window.showSettingsView = function() {
    document.getElementById('mainView').classList.remove('active');
    document.getElementById('settingsView').classList.add('active');
};

window.showMainView = function() {
    document.getElementById('settingsView').classList.remove('active');
    document.getElementById('mainView').classList.add('active');
};

// Load saved settings into the settings form
function loadSettings() {
    const apiProvider = localStorage.getItem('apiProvider') || 'openai';
    const openaiKey = localStorage.getItem('openaiKey') || '';
    const claudeKey = localStorage.getItem('claudeKey') || '';
    const openaiModel = localStorage.getItem('openaiModel') || 'gpt-5.2';
    
    document.getElementById('apiProviderSelect').value = apiProvider;
    document.getElementById('openaiKeyInput').value = openaiKey;
    document.getElementById('claudeKeyInput').value = claudeKey;
    
    if (openaiModel === 'gpt-5.2') {
        document.getElementById('openaiModelSelect').value = 'gpt-5.2';
    } else {
        document.getElementById('openaiModelSelect').value = 'custom';
        document.getElementById('customModelInput').value = openaiModel;
        document.getElementById('customModelDiv').style.display = 'block';
    }
    
    updateProviderFields();
}

// Update provider-specific fields
window.updateProviderFields = function() {
    const provider = document.getElementById('apiProviderSelect').value;
    
    if (provider === 'openai') {
        document.getElementById('openaiSettings').style.display = 'block';
        document.getElementById('claudeSettings').style.display = 'none';
    } else {
        document.getElementById('openaiSettings').style.display = 'none';
        document.getElementById('claudeSettings').style.display = 'block';
    }
};

// Handle model selection
document.addEventListener('DOMContentLoaded', function() {
    const modelSelect = document.getElementById('openaiModelSelect');
    if (modelSelect) {
        modelSelect.addEventListener('change', function() {
            if (this.value === 'custom') {
                document.getElementById('customModelDiv').style.display = 'block';
            } else {
                document.getElementById('customModelDiv').style.display = 'none';
            }
        });
    }
});

// Save settings
window.saveSettings = function() {
    const provider = document.getElementById('apiProviderSelect').value;
    const statusDiv = document.getElementById('settingsStatus');
    
    localStorage.setItem('apiProvider', provider);
    
    if (provider === 'openai') {
        const modelSelect = document.getElementById('openaiModelSelect').value;
        let model = 'gpt-5.2';
        
        if (modelSelect === 'custom') {
            const customModel = document.getElementById('customModelInput').value.trim();
            if (!customModel) {
                showSettingsStatus('Please enter a custom model name', 'error');
                return;
            }
            model = customModel;
        } else {
            model = modelSelect;
        }
        
        const apiKey = document.getElementById('openaiKeyInput').value.trim();
        if (!apiKey) {
            showSettingsStatus('Please enter your OpenAI API key', 'error');
            return;
        }
        
        localStorage.setItem('openaiModel', model);
        localStorage.setItem('openaiKey', apiKey);
        showSettingsStatus('OpenAI settings saved! Model: ' + model, 'success');
        
    } else {
        const apiKey = document.getElementById('claudeKeyInput').value.trim();
        if (!apiKey) {
            showSettingsStatus('Please enter your Claude API key', 'error');
            return;
        }
        
        localStorage.setItem('claudeKey', apiKey);
        showSettingsStatus('Claude settings saved!', 'success');
    }
    
    setTimeout(() => {
        showMainView();
    }, 1500);
};

function showSettingsStatus(message, type) {
    const statusDiv = document.getElementById('settingsStatus');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    statusDiv.style.display = 'block';
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 3000);
}

// Strip signature from email body
function stripSignature(text) {
    const signatureMarker = "Best regards,";
    const index = text.indexOf(signatureMarker);
    
    if (index !== -1) {
        return text.substring(0, index).trim();
    }
    
    return text;
}

// Generate draft using AI
window.generateDraft = async function() {
    const statusDiv = document.getElementById('statusMessage');
    const outputDiv = document.getElementById('draftOutput');
    const actionButtons = document.getElementById('actionButtons');
    const generateBtn = document.getElementById('generateBtn');
    
    // Hide previous results
    outputDiv.style.display = 'none';
    actionButtons.style.display = 'none';
    statusDiv.style.display = 'none';
    
    // Disable button
    generateBtn.disabled = true;
    generateBtn.textContent = 'Generating...';

    try {
        // Get the current email item
        const item = Office.context.mailbox.item;
        
        // Check if subject is empty
        item.subject.getAsync(async (subjectResult) => {
            const hasSubject = subjectResult.status === Office.AsyncResultStatus.Succeeded && 
                              subjectResult.value && 
                              subjectResult.value.trim().length > 0;
            
            // Get email body
            item.body.getAsync(Office.CoercionType.Text, async (bodyResult) => {
                if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    let originalText = bodyResult.value.trim();
                    
                    // Strip signature
                    originalText = stripSignature(originalText);
                    
                    if (!originalText) {
                        showStatus('Please write some notes in the email body first.', 'error');
                        generateBtn.disabled = false;
                        generateBtn.textContent = 'Generate Draft';
                        return;
                    }

                    // Get selected tone
                    const tone = document.getElementById('toneSelect').value;
                    let prompt = tonePrompts[tone];
                    
                    // Add subject generation prompt if subject is empty
                    if (!hasSubject) {
                        prompt += subjectPrompts[tone];
                    }

                    // Get API settings from localStorage
                    const apiProvider = localStorage.getItem('apiProvider') || 'openai';
                    const openaiKey = localStorage.getItem('openaiKey');
                    const claudeKey = localStorage.getItem('claudeKey');

                    // Check if API key exists
                    if (apiProvider === 'openai' && !openaiKey) {
                        showStatus('Please add your OpenAI API key in Settings first.', 'error');
                        generateBtn.disabled = false;
                        generateBtn.textContent = 'Generate Draft';
                        return;
                    }
                    if (apiProvider === 'claude' && !claudeKey) {
                        showStatus('Please add your Claude API key in Settings first.', 'error');
                        generateBtn.disabled = false;
                        generateBtn.textContent = 'Generate Draft';
                        return;
                    }

                    // Call the appropriate API
                    let draftText;
                    if (apiProvider === 'openai') {
                        draftText = await callOpenAI(prompt, originalText, openaiKey);
                    } else {
                        draftText = await callClaude(prompt, originalText, claudeKey);
                    }

                    // Check if response includes a subject line
                    let generatedSubject = null;
                    if (!hasSubject && draftText.startsWith('Subject:')) {
                        const lines = draftText.split('\n');
                        const subjectLine = lines[0];
                        generatedSubject = subjectLine.replace('Subject:', '').trim();
                        
                        // Remove subject line from body
                        draftText = lines.slice(2).join('\n').trim(); // Skip subject and blank line
                        
                        // Set the subject
                        item.subject.setAsync(generatedSubject);
                    }

                    // Display the draft
                    outputDiv.textContent = draftText;
                    outputDiv.style.display = 'block';
                    actionButtons.style.display = 'block';
                    
                    let statusMessage = 'Draft generated successfully!';
                    if (generatedSubject) {
                        statusMessage += ' Subject line added.';
                    }
                    showStatus(statusMessage, 'success');

                } else {
                    showStatus('Error reading email body: ' + bodyResult.error.message, 'error');
                }

                generateBtn.disabled = false;
                generateBtn.textContent = 'Generate Draft';
            });
        });

    } catch (error) {
        showStatus('Error: ' + error.message, 'error');
        generateBtn.disabled = false;
        generateBtn.textContent = 'Generate Draft';
    }
};

// Call OpenAI API
async function callOpenAI(systemPrompt, userText, apiKey) {
    const selectedModel = localStorage.getItem('openaiModel') || 'gpt-5.2';
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            model: selectedModel,
            messages: [
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userText }
            ],
            temperature: 0.7
        })
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'OpenAI API request failed');
    }

    const data = await response.json();
    return data.choices[0].message.content;
}

// Call Claude API
async function callClaude(systemPrompt, userText, apiKey) {
    const selectedModel = localStorage.getItem('claudeModel') || 'claude-3-5-sonnet-20241022';
    
    const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
            model: selectedModel,
            max_tokens: 1024,
            system: systemPrompt,
            messages: [
                { role: 'user', content: userText }
            ]
        })
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Claude API request failed');
    }

    const data = await response.json();
    return data.content[0].text;
}

// Replace entire email body with draft
window.replaceEmail = function() {
    const draftText = document.getElementById('draftOutput').textContent;
    const item = Office.context.mailbox.item;
    
    item.body.setAsync(draftText, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            showStatus('Email replaced successfully!', 'success');
        } else {
            showStatus('Error replacing email: ' + result.error.message, 'error');
        }
    });
};

// Insert draft below existing text
window.insertBelow = function() {
    const draftText = document.getElementById('draftOutput').textContent;
    const item = Office.context.mailbox.item;
    
    item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const combined = result.value + '\n\n---\n\n' + draftText;
            item.body.setAsync(combined, { coercionType: Office.CoercionType.Text }, (result2) => {
                if (result2.status === Office.AsyncResultStatus.Succeeded) {
                    showStatus('Draft inserted below!', 'success');
                } else {
                    showStatus('Error inserting draft: ' + result2.error.message, 'error');
                }
            });
        }
    });
};

// Show status message
function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    statusDiv.style.display = 'block';
}
