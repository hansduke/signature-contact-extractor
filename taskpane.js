let currentMessage;
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('extractBtn').onclick = extractSignature;
        document.getElementById('createBtn').onclick = createContact;
        loadCurrentMessage();
    }
});
function showStatus(message, type = 'info') {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    if (type === 'success') {
        setTimeout(() => { statusDiv.className = 'status'; }, 5000);
    }
}
function loadCurrentMessage() {
    const item = Office.context.mailbox.item;
    if (item) {
        currentMessage = item;
        showStatus('Email loaded. Click "Auto-Extract from Signature" to extract contact info.', 'info');
    }
}
function extractSignature() {
    if (!currentMessage) {
        showStatus('No email loaded. Please open an email first.', 'error');
        return;
    }
    currentMessage.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const htmlBody = result.value;
            const signatureData = parseSignatureAdvanced(htmlBody);
            if (signatureData && Object.keys(signatureData).length > 0) {
                populateForm(signatureData);
                showStatus('Signature extracted successfully!', 'success');
            } else {
                showStatus('Could not extract signature. Please fill in manually.', 'error');
            }
        } else {
            showStatus('Error: ' + result.error.message, 'error');
        }
    });
}
function parseSignatureAdvanced(htmlBody) {
    const data = {};
    const plainText = htmlToPlainText(htmlBody);
    const lines = plainText.split('\\n').map(l => l.trim()).filter(l => l && l.length > 0);
    if (lines.length === 0) return data;
    const emailMatch = plainText.match(/([a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})/i);
    if (emailMatch) data.email = emailMatch[1];
    const phonePattern = /\\(?(\\d{3})\\)?[-.\\s]?(\\d{3})[-.\\s]?(\\d{4})/g;
    const phones = [];
    let phoneMatch;
    while ((phoneMatch = phonePattern.exec(plainText)) !== null) {
        if (!phones.includes(phoneMatch[0])) phones.push(phoneMatch[0]);
    }
    if (phones.length > 0) data.mobilePhone = phones[0];
    if (phones.length > 1) data.workPhone = phones[1];
    const nameMatch = lines[0].match(/^([A-Z][a-z]+(?:\\s+[A-Z][a-z]+)*)\\s*$/);
    if (nameMatch) {
        const fullName = nameMatch[1];
        const nameParts = fullName.split(/\\s+/);
        data.firstName = nameParts[0];
        data.lastName = nameParts.slice(1).join(' ');
    }
    const titleKeywords = ['Director', 'Manager', 'Officer', 'Specialist', 'Chief', 'Deputy', 'Assistant'];
    for (let i = 0; i < lines.length; i++) {
        const hasTitle = titleKeywords.some(kw => lines[i].toLowerCase().includes(kw.toLowerCase()));
        if (hasTitle && lines[i].length < 100 && !lines[i].includes('@')) {
            data.jobTitle = lines[i];
            break;
        }
    }
    const deptKeywords = ['Division', 'Department', 'Office', 'Bureau'];
    for (let i = 0; i < lines.length; i++) {
        const hasDept = deptKeywords.some(kw => lines[i].toLowerCase().includes(kw.toLowerCase()));
        if (hasDept && !lines[i].includes('@')) {
            data.department = lines[i];
            break;
        }
    }
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.length > 15 && !line.includes('@') && !line.match(/\\d{3}/) && !line.toLowerCase().includes('confidential')) {
            data.company = line;
            break;
        }
    }
    return data;
}
function htmlToPlainText(html) {
    let text = html.replace(/<script\\b[^<]*(?:(?!<\\/script>)<[^<]*)*<\\/script>/gi, '');
    text = text.replace(/<style\\b[^<]*(?:(?!<\\/style>)<[^<]*)*<\\/style>/gi, '');
    text = text.replace(/<br\\s*\\/?>/gi, '\\n');
    text = text.replace(/<\\/(p|div|li)>/gi, '\\n');
    text = text.replace(/<[^>]*>/g, '');
    text = text.replace(/&amp;/g, '&');
    text = text.replace(/&lt;/g, '<');
    text = text.replace(/&gt;/g, '>');
    text = text.replace(/&quot;/g, '"');
    text = text.replace(/&#39;/g, "'");
    text = text.replace(/&nbsp;/g, ' ');
    text = text.replace(/\\s+/g, ' ');
    text = text.replace(/\\n\\s+/g, '\\n');
    return text;
}
function populateForm(data) {
    if (data.firstName) document.getElementById('firstName').value = data.firstName;
    if (data.lastName) document.getElementById('lastName').value = data.lastName;
    if (data.email) document.getElementById('email').value = data.email;
    if (data.mobilePhone) document.getElementById('mobilePhone').value = data.mobilePhone;
    if (data.workPhone) document.getElementById('workPhone').value = data.workPhone;
    if (data.jobTitle) document.getElementById('jobTitle').value = data.jobTitle;
    if (data.department) document.getElementById('department').value = data.department;
    if (data.company) document.getElementById('company').value = data.company;
}
async function createContact() {
    const firstName = document.getElementById('firstName').value.trim();
    const lastName = document.getElementById('lastName').value.trim();
    const email = document.getElementById('email').value.trim();
    if (!firstName || !lastName || !email) {
        showStatus('Please fill in First Name, Last Name, and Email.', 'error');
        return;
    }
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            showStatus('Error: ' + result.error.message, 'error');
            return;
        }
        const token = result.value;
        const contactData = {
            givenName: firstName,
            surname: lastName,
            emailAddresses: [{ address: email, name: firstName + ' ' + lastName }],
            businessPhones: document.getElementById('workPhone').value.trim() ? [document.getElementById('workPhone').value.trim()] : [],
            mobilePhone: document.getElementById('mobilePhone').value.trim() || null,
            jobTitle: document.getElementById('jobTitle').value.trim() || null,
            department: document.getElementById('department').value.trim() || null,
            companyName: document.getElementById('company').value.trim() || null
        };
        Object.keys(contactData).forEach(key => contactData[key] === null && delete contactData[key]);
        try {
            const response = await fetch('https://graph.microsoft.com/v1.0/me/contacts', {
                method: 'POST',
                headers: {
                    'Authorization': 'Bearer ' + token,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(contactData)
            });
            if (response.ok) {
                showStatus('Contact created successfully!', 'success');
                document.getElementById('firstName').value = '';
                document.getElementById('lastName').value = '';
                document.getElementById('email').value = '';
                document.getElementById('mobilePhone').value = '';
                document.getElementById('workPhone').value = '';
                document.getElementById('jobTitle').value = '';
                document.getElementById('department').value = '';
                document.getElementById('company').value = '';
            } else {
                const errorData = await response.json();
                showStatus('Error: ' + (errorData.error?.message || response.statusText), 'error');
            }
        } catch (error) {
            showStatus('Error: ' + error.message, 'error');
        }
    });
}
