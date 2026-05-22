// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Email Drafter loaded successfully");
        loadSettings();
        restoreLastTone();
        document.getElementById('toneSelect').addEventListener('change', (e) => {
            localStorage.setItem('lastTone', e.target.value);
        });
    }
});

function restoreLastTone() {
    const saved = localStorage.getItem('lastTone');
    if (!saved) return;
    const select = document.getElementById('toneSelect');
    if ([...select.options].some(o => o.value === saved)) {
        select.value = saved;
    }
}

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

// Collect the text content of `block` that appears in document order
// before (targetNode, targetOffset). Used to detect whether `block` is
// a tight signature paragraph (no notes content before the sig) or a
// big container that holds both notes and sig (notes text is found).
function textBeforeInBlock(doc, block, targetNode, targetOffset) {
    const tw = doc.createTreeWalker(block, NodeFilter.SHOW_TEXT, null);
    let acc = '';
    let n;
    while ((n = tw.nextNode())) {
        if (n === targetNode) {
            acc += n.nodeValue.substring(0, targetOffset);
            break;
        }
        acc += n.nodeValue;
    }
    return acc;
}

// Parse the Outlook body as a DOM, locate the "Best regards," text node,
// and delete everything from the start of <body> to that text position
// via a Range. Range.deleteContents handles structural cuts safely
// without breaking tag pairs, so the chain (anything after the sig in
// document order) is left untouched.
//
// Detects two structural cases:
//   A) Each paragraph is its own <p>/<div> block at body level. The sig
//      block is a sibling of the notes blocks. Insert paragraph clones
//      of the sig block (preserving its class/style) as previous
//      siblings of the sig block.
//   B) Notes and sig live inside a single styled container element with
//      <br>-separated text. The sig text node is a child of that
//      container. Insert bare <p> elements inside the container, just
//      before the sig text node; they inherit font/size from the
//      container's inline style.
//
// Either way the inserted blocks become real paragraph elements, so the
// visual gap between paragraphs is a "proper" paragraph break, not
// stacked <br>s.
function surgicalBodyReplace(originalHtml, draftText) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(originalHtml, 'text/html');
    const body = doc.body;

    let sigText = null;
    let sigOff = -1;
    const tw = doc.createTreeWalker(body, NodeFilter.SHOW_TEXT, null);
    while (tw.nextNode()) {
        const idx = tw.currentNode.nodeValue.indexOf('Best regards,');
        if (idx !== -1) {
            sigText = tw.currentNode;
            sigOff = idx;
            break;
        }
    }

    if (!sigText) {
        // No sig anchor — wipe the body and write the draft as bare <p>s.
        while (body.firstChild) body.removeChild(body.firstChild);
        for (const para of splitParagraphs(draftText)) {
            body.appendChild(buildParagraph(doc, doc.createElement('p'), para));
        }
        return { html: doc.documentElement.outerHTML, signaturePreserved: false };
    }

    // Closest <p>/<div> ancestor of the sig text — the "sig block."
    let sigBlock = sigText.parentElement;
    while (sigBlock && sigBlock !== body && sigBlock.tagName !== 'P' && sigBlock.tagName !== 'DIV') {
        sigBlock = sigBlock.parentElement;
    }
    if (!sigBlock || sigBlock === body) sigBlock = sigText.parentElement;

    const notesTextInBlock = textBeforeInBlock(doc, sigBlock, sigText, sigOff).trim();
    const isBigContainer = notesTextInBlock.length > 0;

    let templateProto;
    let insertContainer;
    let insertBeforeRef;
    if (isBigContainer) {
        // Case B — paragraphs go inside the styled container, before sig text.
        templateProto = doc.createElement('p');
        insertContainer = sigBlock;
        insertBeforeRef = sigText;
    } else {
        // Case A — clone sig block's tag/attrs (font styling lives there),
        // insert clones as previous siblings of the sig block.
        templateProto = sigBlock.cloneNode(false);
        insertContainer = sigBlock.parentElement;
        insertBeforeRef = sigBlock;
    }

    const range = doc.createRange();
    range.setStart(body, 0);
    range.setEnd(sigText, sigOff);
    range.deleteContents();

    for (const para of splitParagraphs(draftText)) {
        const block = templateProto.cloneNode(false);
        block.innerHTML = '';
        buildParagraph(doc, block, para);
        insertContainer.insertBefore(block, insertBeforeRef);
    }

    return { html: doc.documentElement.outerHTML, signaturePreserved: true };
}

function splitParagraphs(text) {
    return text
        .split(/\n\s*\n/)
        .map(p => p.replace(/^\s+|\s+$/g, ''))
        .filter(p => p.length > 0);
}

// Fill `block` with text + <br>s for single-newline line breaks.
function buildParagraph(doc, block, para) {
    const lines = para.split('\n');
    lines.forEach((line, i) => {
        if (i > 0) block.appendChild(doc.createElement('br'));
        block.appendChild(doc.createTextNode(line));
    });
    return block;
}

// Replace the user's notes in the email body with the draft, keeping the
// signature and quoted chain intact.
window.replaceBody = function() {
    const draftText = document.getElementById('draftTextarea').value;
    const item = Office.context.mailbox.item;

    item.body.getAsync(Office.CoercionType.Html, (getResult) => {
        if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
            showStatus('Error reading body: ' + getResult.error.message, 'error');
            return;
        }

        let result;
        try {
            result = surgicalBodyReplace(getResult.value, draftText);
        } catch (e) {
            showStatus('Error rebuilding body: ' + e.message, 'error');
            return;
        }

        item.body.setAsync(result.html, { coercionType: Office.CoercionType.Html }, (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                const sigNote = result.signaturePreserved ? ' (signature + chain preserved)' : ' (no signature found)';
                showStatus('Email body replaced' + sigNote + '.', 'success');
            } else {
                showStatus('Error: ' + setResult.error.message, 'error');
            }
        });
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
