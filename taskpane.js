// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Email Drafter loaded successfully");
    }
});

// Tone prompts for different email styles
const tonePrompts = {
    professional: "Rewrite the following notes as a polished, professional business email. Use formal language, proper structure, and maintain a respectful tone. Keep it clear and concise.",
    friendly: "Rewrite the following notes as a warm, friendly professional email. Be approachable and personable while maintaining professionalism. Use a conversational but business-appropriate tone.",
    brief: "Rewrite the following notes as a brief, direct email. Get straight to the point. Use short sentences and minimal pleasantries. Be clear and action-oriented.",
    casual: "Rewrite the following notes as a casual, relaxed email. Use a conversational tone as if writing to a colleague or familiar contact. Keep it friendly and informal.",
    diplomatic: "Rewrite the following notes as a diplomatic, apologetic email. Handle the situation delicately, acknowledge any issues, and maintain a professional yet empathetic tone. Focus on solutions and positive resolution."
};

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
        // Get the current email body
        const item = Office.context.mailbox.item;
        
        item.body.getAsync(Office.CoercionType.Text, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const originalText = result.value.trim();
                
                if (!originalText) {
                    showStatus('Please write some notes in the email body first.', 'error');
                    generateBtn.disabled = false;
                    generateBtn.textContent = 'Generate Draft';
                    return;
                }

                // Get selected tone
                const tone = document.getElementById('toneSelect').value;
                const prompt = tonePrompts[tone];

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

                // Display the draft
                outputDiv.textContent = draftText;
                outputDiv.style.display = 'block';
                actionButtons.style.display = 'block';
                showStatus('Draft generated successfully!', 'success');

            } else {
                showStatus('Error reading email body: ' + result.error.message, 'error');
            }

            generateBtn.disabled = false;
            generateBtn.textContent = 'Generate Draft';
        });

    } catch (error) {
        showStatus('Error: ' + error.message, 'error');
        generateBtn.disabled = false;
        generateBtn.textContent = 'Generate Draft';
    }
};

// Call OpenAI API
async function callOpenAI(systemPrompt, userText, apiKey) {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            model: 'gpt-4o-mini',
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
    const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
            model: 'claude-3-5-sonnet-20241022',
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

// Show settings (placeholder for now)
window.showSettings = function() {
    const apiProvider = localStorage.getItem('apiProvider') || 'openai';
    const openaiKey = localStorage.getItem('openaiKey') || '';
    const claudeKey = localStorage.getItem('claudeKey') || '';
    
    const newProvider = prompt('API Provider (enter "openai" or "claude"):', apiProvider);
    if (newProvider && (newProvider === 'openai' || newProvider === 'claude')) {
        localStorage.setItem('apiProvider', newProvider);
    }
    
    if (newProvider === 'openai' || apiProvider === 'openai') {
        const newOpenAIKey = prompt('Enter your OpenAI API Key:', openaiKey);
        if (newOpenAIKey) {
            localStorage.setItem('openaiKey', newOpenAIKey);
            alert('OpenAI API key saved!');
        }
    }
    
    if (newProvider === 'claude' || apiProvider === 'claude') {
        const newClaudeKey = prompt('Enter your Claude API Key:', claudeKey);
        if (newClaudeKey) {
            localStorage.setItem('claudeKey', newClaudeKey);
            alert('Claude API key saved!');
        }
    }
};
