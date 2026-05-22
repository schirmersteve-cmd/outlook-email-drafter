// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Email Drafter loaded successfully");
        loadSettings();
    }
});

const MODEL = 'claude-sonnet-4-6';
const MAX_TOKENS = 1024;

// Tone prompts for different email styles
const tonePrompts = {
    professional: "Rewrite the following notes as a polished, professional business email. Use formal language, proper structure, and a respectful tone. Keep it clear and concise.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice. Do not add commitments, promises, or technical claims.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    friendly: "Rewrite the following notes as a warm, friendly professional email. Be approachable and personable while maintaining professionalism. Use a conversational but business-appropriate tone.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    casual: "Rewrite the following notes as a casual, relaxed email. Use a conversational tone as if writing to a colleague or familiar contact. Keep it friendly and informal while remaining appropriate for business communication.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    brief: "Rewrite the following notes as a brief, direct email. Get straight to the point. Use short sentences and minimal pleasantries. Be clear and action-oriented without sounding abrupt or curt.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    diplomatic: "Rewrite the following notes as a diplomatic, tactful email. Handle the situation delicately, acknowledge concerns when appropriate, and maintain a professional and empathetic tone. Focus on solutions and constructive next steps.\n\nPreserve original meaning and factual content. Do not invent details. Maintain the sender's voice. Do not admit fault or liability unless explicitly stated in the notes.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    cleanup: "Clean up the following notes or draft email.\n\nFix grammar and spelling.\nRemove unnecessary fluff.\nImprove readability and flow.\nNormalize formatting.\nKeep the message concise.\n\nDo not change the intent, structure, or meaning of the content. Do not introduce new phrasing that alters emphasis or tone. Preserve original factual content and the sender's voice.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    sales: "Rewrite the following notes as a professional, customer-facing business email.\n\nImprove clarity, persuasion, and engagement while maintaining credibility and technical accuracy. Keep the tone confident, practical, and relationship-driven, not marketing-oriented or promotional.\n\nPreserve original meaning and factual content. Do not invent specifications, performance claims, or commitments. Maintain the sender's voice.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes.",

    myvoice: "Convert the following rough notes, shorthand, or bullet points into a complete professional email written in the sender's natural voice.\n\nGuidelines:\n• Write in a style that is professional but conversational.\n• Be direct, clear, and relationship-focused.\n• Avoid corporate jargon, marketing language, or overly formal phrasing.\n• Maintain technical credibility without sounding like a brochure.\n• Keep the message natural and easy to read.\n\nContent Rules:\n• Preserve original meaning and factual content.\n• Do not invent details, specifications, or commitments.\n• If notes are vague, keep language appropriately general and safe.\n\nStructure:\n• Organize the message into a logical email flow.\n• Add transitions and readability improvements where needed.\n• Keep length appropriate to the content. Do not over-expand.\n\nDo not add any closing lines (like 'Thanks', 'Best regards', etc.) - the sender has their own signature.\n\nDo not use em dashes."
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
    const claudeKey = localStorage.getItem('claudeKey') || '';
    document.getElementById('claudeKeyInput').value = claudeKey;
}

// Save settings
window.saveSettings = function() {
    const apiKey = document.getElementById('claudeKeyInput').value.trim();
    if (!apiKey) {
        showSettingsStatus('Please enter your Anthropic API key', 'error');
        return;
    }

    localStorage.setItem('claudeKey', apiKey);
    showSettingsStatus('Settings saved!', 'success');

    setTimeout(showMainView, 1500);
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

// Strip signature from email body (plain text)
function stripSignature(text) {
    const signatureMarker = "Best regards,";
    const index = text.indexOf(signatureMarker);

    if (index !== -1) {
        return text.substring(0, index).trim();
    }

    return text;
}

// Generate draft using Claude
window.generateDraft = async function() {
    const statusDiv = document.getElementById('statusMessage');
    const outputDiv = document.getElementById('draftOutput');
    const draftTextarea = document.getElementById('draftTextarea');
    const actionButtons = document.getElementById('actionButtons');
    const generateBtn = document.getElementById('generateBtn');

    outputDiv.style.display = 'none';
    actionButtons.style.display = 'none';
    statusDiv.style.display = 'none';

    generateBtn.disabled = true;
    generateBtn.textContent = 'Generating...';

    try {
        const item = Office.context.mailbox.item;

        item.subject.getAsync(async (subjectResult) => {
            const hasSubject = subjectResult.status === Office.AsyncResultStatus.Succeeded &&
                              subjectResult.value &&
                              subjectResult.value.trim().length > 0;

            item.body.getAsync(Office.CoercionType.Text, async (bodyResult) => {
                if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    let originalText = bodyResult.value.trim();
                    originalText = stripSignature(originalText);

                    if (!originalText) {
                        showStatus('Please write some notes in the email body first.', 'error');
                        generateBtn.disabled = false;
                        generateBtn.textContent = 'Generate Draft';
                        return;
                    }

                    const tone = document.getElementById('toneSelect').value;
                    let prompt = tonePrompts[tone];

                    if (!hasSubject) {
                        prompt += subjectPrompts[tone];
                    }

                    const claudeKey = localStorage.getItem('claudeKey');
                    if (!claudeKey) {
                        showStatus('Please add your Anthropic API key in Settings first.', 'error');
                        generateBtn.disabled = false;
                        generateBtn.textContent = 'Generate Draft';
                        return;
                    }

                    let draftText;
                    try {
                        draftText = await callClaude(prompt, originalText, claudeKey);
                    } catch (apiErr) {
                        showStatus('Error: ' + apiErr.message, 'error');
                        generateBtn.disabled = false;
                        generateBtn.textContent = 'Generate Draft';
                        return;
                    }

                    let generatedSubject = null;
                    if (!hasSubject && draftText.startsWith('Subject:')) {
                        const lines = draftText.split('\n');
                        const subjectLine = lines[0];
                        generatedSubject = subjectLine.replace('Subject:', '').trim();
                        draftText = lines.slice(2).join('\n').trim();
                        item.subject.setAsync(generatedSubject);
                    }

                    draftTextarea.value = draftText;
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

// Call Anthropic API
// anthropic-dangerous-direct-browser-access is required for browser-origin calls;
// Anthropic blocks direct browser fetches by default to discourage key leakage.
// Acceptable here because the key lives in localStorage on Steve's machines only.
async function callClaude(systemPrompt, userText, apiKey) {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01',
            'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({
            model: MODEL,
            max_tokens: MAX_TOKENS,
            system: systemPrompt,
            messages: [
                { role: 'user', content: userText }
            ]
        })
    });

    if (!response.ok) {
        const error = await response.json().catch(() => ({}));
        throw new Error(error.error?.message || `Anthropic API request failed (${response.status})`);
    }

    const data = await response.json();
    return data.content[0].text;
}

// Replace selected text with draft
window.replaceSelection = function() {
    const draftText = document.getElementById('draftTextarea').value;
    const item = Office.context.mailbox.item;

    item.body.setSelectedDataAsync(draftText, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            showStatus('Selection replaced successfully!', 'success');
        } else {
            showStatus('Error: ' + result.error.message + '. Make sure you have text selected in the email.', 'error');
        }
    });
};

// Copy draft to clipboard
window.copyToClipboard = function() {
    const draftText = document.getElementById('draftTextarea').value;

    navigator.clipboard.writeText(draftText).then(() => {
        showStatus('Draft copied to clipboard!', 'success');
    }).catch(err => {
        showStatus('Error copying to clipboard: ' + err.message, 'error');
    });
};

// Show status message
function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    statusDiv.style.display = 'block';
}
