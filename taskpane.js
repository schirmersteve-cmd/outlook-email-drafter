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

const MODEL = 'claude-sonnet-5';
const MAX_TOKENS = 1024;

// The taskpane's "Currently using:" label is derived from MODEL rather than
// written out in the HTML, so it cannot drift out of step with the model
// actually being called (it sat on "Claude Sonnet 4.6" after the move to
// Sonnet 5). Trailing numeric segments rejoin with a dot:
// claude-sonnet-5 -> "Claude Sonnet 5", claude-opus-4-8 -> "Claude Opus 4.8".
function modelDisplayName(id) {
    const words = [];
    const nums = [];
    for (const part of String(id).split('-')) {
        if (/^\d+$/.test(part)) nums.push(part);
        else if (part) words.push(part[0].toUpperCase() + part.slice(1));
    }
    const name = words.join(' ');
    return nums.length ? `${name} ${nums.join('.')}` : name;
}

// taskpane.js is loaded at the end of <body>, so the element already exists.
(function showModelName() {
    const el = document.getElementById('modelName');
    if (el) el.textContent = modelDisplayName(MODEL);
})();

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

// Appended when the compose item already has a subject (replies, forwards, or
// a new mail Steve already titled). Without it the model volunteers a
// "Subject: ..." line anyway and it ends up pasted into the body.
const noSubjectInstruction = "\n\nThis email already has a subject line. Return the email body only. Do not include a 'Subject:' line or any heading before the body.";

// Pull a leading "Subject: ..." line off the model's reply.
// Runs on every response, not just the no-subject case: the instruction above
// reduces stray subject lines but does not guarantee their absence, and a
// leftover line is worse in the body than discarded. The caller decides
// whether the value is worth writing to item.subject.
function extractSubjectLine(text) {
    const lines = text.split('\n');
    let i = 0;
    while (i < lines.length && !lines[i].trim()) i++;

    const match = lines[i] ? lines[i].trim().match(/^\*{0,2}subject\*{0,2}\s*:\*{0,2}\s*(.*)$/i) : null;
    if (!match) return { subject: null, body: text };

    const rest = lines.slice(i + 1);
    while (rest.length && !rest[0].trim()) rest.shift();

    return {
        subject: match[1].replace(/\*+$/, '').trim(),
        body: rest.join('\n').trim()
    };
}

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

                    prompt += hasSubject ? noSubjectInstruction : subjectPrompts[tone];

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

                    const parsed = extractSubjectLine(draftText);
                    let generatedSubject = null;
                    if (parsed.subject) {
                        // Strip it either way; only write it when the item has
                        // no subject of its own, so a reply's "RE: ..." stands.
                        draftText = parsed.body;
                        if (!hasSubject) {
                            generatedSubject = parsed.subject;
                            item.subject.setAsync(generatedSubject);
                        }
                    }

                    draftTextarea.value = draftText;
                    // 'flex', not 'block': an inline block here would beat the
                    // stylesheet and collapse the fill-the-pane layout.
                    outputDiv.style.display = 'flex';
                    actionButtons.style.display = 'flex';

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
            // Sonnet 5 turns adaptive thinking ON when this field is omitted
            // (Sonnet 4.6 ran it off). max_tokens caps thinking + reply
            // together, so leaving it on would eat the 1024-token budget and
            // truncate drafts. Disabled keeps behavior identical to 4.6 and
            // keeps the taskpane snappy. Remove this line (and raise
            // MAX_TOKENS) if better drafts are worth the extra latency.
            thinking: { type: 'disabled' },
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

// Outlook's current compose default (what p.MsoNormal resolves to). Only used
// when the body gives us nothing to copy from.
const DEFAULT_PARA_FONT = 'Aptos, sans-serif';
const DEFAULT_PARA_SIZE = '11.0pt';

// The font actually in force at `el`, read off inline styles up the tree.
//
// Why not just set class="MsoNormal" and let the stylesheet do it: Office.js
// hands back a body *fragment*, so the <style> block that defines p.MsoNormal
// generally does not come with it. Without that rule the class is inert and
// Word falls back to Times New Roman 12pt with default paragraph margins —
// wrong font AND wrong spacing. Stamping the values inline is what survives
// the getAsync/setAsync round trip.
// The p.MsoNormal rule out of the message's own <style> block, when one came
// through. This is the message's real body font, so it beats any constant we
// could hardcode — Outlook's default changed from Calibri to Aptos not long
// ago, and guessing wrong puts our paragraphs out of step with the signature.
function msoNormalFont(html) {
    const rule = /p\.MsoNormal[^{]*\{([^}]*)\}/i.exec(html || '');
    if (!rule) return {};
    const fam = /(?:^|[;\s])font-family\s*:\s*([^;]+)/i.exec(rule[1]);
    const size = /(?:^|[;\s])font-size\s*:\s*([^;]+)/i.exec(rule[1]);
    return {
        family: fam ? fam[1].trim() : '',
        size: size ? size[1].trim() : ''
    };
}

// The font actually in force at `el`. Inline styles on the signature's own
// ancestors win, then the stylesheet's p.MsoNormal rule, then the constants.
function inheritedFont(el, sheetFont) {
    let family = '';
    let size = '';
    let n = el;
    while (n && n.nodeType === 1) {
        if (n.style) {
            if (!family && n.style.fontFamily) family = n.style.fontFamily;
            if (!size && n.style.fontSize) size = n.style.fontSize;
        }
        if (family && size) break;
        n = n.parentElement;
    }
    const s = sheetFont || {};
    return {
        family: family || s.family || DEFAULT_PARA_FONT,
        size: size || s.size || DEFAULT_PARA_SIZE
    };
}

// Make a paragraph render identically to the signature: same font, and
// margin:0 to match p.MsoNormal. Keep the class too when the document uses it,
// so the styled and inline paths agree if the stylesheet IS present.
function stampParaStyle(doc, p, font) {
    if (!p.className && doc.querySelector('.MsoNormal')) p.className = 'MsoNormal';
    p.style.fontFamily = font.family;
    p.style.fontSize = font.size;
    p.style.marginTop = '0';
    p.style.marginBottom = '0';
    return p;
}

// A real empty paragraph. With margin:0 this is what produces the visible gap
// between paragraphs — the same way Outlook does it when you press Enter twice.
// Margins can't do the job here: matching the signature means margin:0, and
// leaving the default margins in place is what made Case B's spacing too loose
// while Case A's was too tight.
function blankPara(doc, font) {
    const p = stampParaStyle(doc, doc.createElement('p'), font);
    p.appendChild(doc.createTextNode(' '));
    return p;
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
//      container. Insert <p> elements inside the container, just before
//      the sig text node.
//
// Either way the inserted blocks become real paragraph elements, so the
// visual gap between paragraphs is a "proper" paragraph break, not
// stacked <br>s. Every inserted paragraph gets its font and margins
// stamped inline by stampParaStyle, and paragraphs are separated by real
// empty paragraphs rather than by margins.
//
// The two cases used to disagree, which is why this kept not converging:
// Case A cloned the sig's p.MsoNormal (margin:0) and came out with NO gap
// between paragraphs, while Case B used a bare <p> that picked up Word's
// default Times 12pt AND ~1em margins — too loose, and the wrong font.
// Fixing one made the other worse. Both now go through the same styling.
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
        // No sig anchor — wipe the body and write the draft out fresh.
        const sheetFont = msoNormalFont(originalHtml);
        const font = inheritedFont(body, sheetFont);
        while (body.firstChild) body.removeChild(body.firstChild);
        splitParagraphs(draftText).forEach((para, i) => {
            if (i > 0) body.appendChild(blankPara(doc, font));
            const p = stampParaStyle(doc, doc.createElement('p'), font);
            body.appendChild(buildParagraph(doc, p, para));
        });
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

    // Read the font before the delete below, while the sig's ancestor chain is
    // still fully intact.
    const font = inheritedFont(isBigContainer ? sigBlock : sigText.parentElement,
                               msoNormalFont(originalHtml));

    const range = doc.createRange();
    range.setStart(body, 0);
    range.setEnd(sigText, sigOff);
    range.deleteContents();

    splitParagraphs(draftText).forEach((para, i) => {
        if (i > 0) insertContainer.insertBefore(blankPara(doc, font), insertBeforeRef);
        const block = stampParaStyle(doc, templateProto.cloneNode(false), font);
        block.innerHTML = '';
        buildParagraph(doc, block, para);
        insertContainer.insertBefore(block, insertBeforeRef);
    });

    // The range delete above removed whatever blank line separated the notes
    // from the signature, so put one back.
    insertContainer.insertBefore(blankPara(doc, font), insertBeforeRef);

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
