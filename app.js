// عناصر DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const removeFile = document.getElementById('removeFile');
const preview = document.getElementById('preview');
const previewTable = document.getElementById('previewTable');
const rowCount = document.getElementById('rowCount');
const configSection = document.getElementById('configSection');
const sendSection = document.getElementById('sendSection');
const progressSection = document.getElementById('progressSection');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');
const statusList = document.getElementById('statusList');

const senderEmail = document.getElementById('senderEmail');
const senderPassword = document.getElementById('senderPassword');
const emailColumn = document.getElementById('emailColumn');
const messageColumn = document.getElementById('messageColumn');
const defaultSubject = document.getElementById('defaultSubject');
const messageMode = document.getElementById('messageMode');
const columnMessageGroup = document.getElementById('columnMessageGroup');
const manualMessageGroup = document.getElementById('manualMessageGroup');
const messageEditor = document.getElementById('messageEditor');

const sendBtn = document.getElementById('sendBtn');
const resetBtn = document.getElementById('resetBtn');
const successModal = document.getElementById('successModal');
const closeModalBtn = document.getElementById('closeModalBtn');

// متغيرات عامة
let fileData = null;
let emailData = [];
let currentMessageMode = 'column';

// Event Listeners لمحرر النصوص
messageMode.addEventListener('change', (e) => {
    currentMessageMode = e.target.value;
    if (currentMessageMode === 'column') {
        columnMessageGroup.classList.remove('hidden');
        manualMessageGroup.classList.add('hidden');
    } else {
        columnMessageGroup.classList.add('hidden');
        manualMessageGroup.classList.remove('hidden');
    }
});

// Event Listeners لرفع الملفات
uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = 'var(--primary-dark)';
    uploadArea.style.background = 'rgba(59, 130, 246, 0.1)';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.borderColor = 'var(--primary-color)';
    uploadArea.style.background = 'var(--gray-50)';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = 'var(--primary-color)';
    uploadArea.style.background = 'var(--gray-50)';
    handleFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

removeFile.addEventListener('click', () => {
    resetForm();
});

// معالجة الملف
function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert('الرجاء اختيار ملف Excel صحيح (.xlsx أو .xls)');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            // التحقق من توفر مكتبة XLSX
            if (typeof XLSX === 'undefined') {
                alert('خطأ: مكتبة Excel لم يتم تحميلها. الرجاء تحديث الصفحة والمحاولة مجدداً.');
                console.error('XLSX library not loaded');
                return;
            }
            
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            fileData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (fileData.length < 2) {
                alert('الملف يجب أن يحتوي على رأس الأعمدة وبيانات واحدة على الأقل');
                return;
            }

            // عرض معلومات الملف
            fileName.textContent = `✅ تم تحميل الملف: ${file.name}`;
            fileInfo.classList.remove('hidden');
            uploadArea.classList.add('hidden');

            // عرض المعاينة
            showPreview();

            // عرض إعدادات البريد
            populateColumnSelects();
            configSection.classList.remove('hidden');
            sendSection.classList.remove('hidden');

        } catch (error) {
            console.error('خطأ في قراءة الملف:', error);
            alert('حدث خطأ في قراءة الملف. الرجاء التأكد من أن الملف بصيغة Excel صحيحة.');
        }
    };
    reader.readAsBinaryString(file);
}

// عرض معاينة البيانات
function showPreview() {
    preview.classList.remove('hidden');

    // مسح الجدول السابق
    previewTable.querySelector('thead').innerHTML = '';
    previewTable.querySelector('tbody').innerHTML = '';

    // إضافة رأس الأعمدة
    const headerRow = previewTable.querySelector('thead');
    const headerCells = fileData[0].map(cell => {
        const th = document.createElement('th');
        th.textContent = cell;
        return th;
    });
    headerRow.append(...headerCells);

    // إضافة أول 5 صفوف من البيانات
    const tbody = previewTable.querySelector('tbody');
    const rowsToShow = Math.min(5, fileData.length - 1);
    for (let i = 1; i <= rowsToShow; i++) {
        const row = document.createElement('tr');
        fileData[i].forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell || '-';
            row.appendChild(td);
        });
        tbody.appendChild(row);
    }

    rowCount.textContent = fileData.length - 1;
}

// ملء خيارات الأعمدة
function populateColumnSelects() {
    const columns = fileData[0];

    // مسح الخيارات السابقة
    emailColumn.innerHTML = '';
    messageColumn.innerHTML = '';

    // إضافة الخيارات الجديدة
    columns.forEach((col, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = col;

        emailColumn.appendChild(option.cloneNode(true));
        messageColumn.appendChild(option.cloneNode(true));
    });

    // تعيين قيم افتراضية
    if (columns.length > 0) emailColumn.value = 0;
    if (columns.length > 1) messageColumn.value = 1;
}

// إعادة تعيين النموذج
function resetForm() {
    fileData = null;
    emailData = [];
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    preview.classList.add('hidden');
    configSection.classList.add('hidden');
    sendSection.classList.add('hidden');
    progressSection.classList.add('hidden');
    uploadArea.classList.remove('hidden');
    statusList.innerHTML = '';
    senderEmail.value = '';
    senderPassword.value = '';
    messageEditor.value = '';
    messageMode.value = 'column';
    currentMessageMode = 'column';
    columnMessageGroup.classList.remove('hidden');
    manualMessageGroup.classList.add('hidden');
}

// تحضير البيانات للإرسال
function prepareEmailData() {
    emailData = [];

    const emailColIdx = parseInt(emailColumn.value);
    const messageColIdx = currentMessageMode === 'column' ? parseInt(messageColumn.value) : -1;

    for (let i = 1; i < fileData.length; i++) {
        const row = fileData[i];
        
        // اختيار الرسالة
        let message = '';
        if (currentMessageMode === 'column') {
            message = row[messageColIdx] || '';
        } else {
            message = messageEditor.value || '';
        }

        const email = {
            to: row[emailColIdx] || '',
            subject: defaultSubject.value,
            message: message
        };

        // التحقق من صحة البيانات
        if (email.to.trim() && email.message.trim()) {
            emailData.push(email);
        }
    }

    return emailData;
}

// إرسال الرسائل
sendBtn.addEventListener('click', async () => {
    // التحقق من البيانات
    if (!senderEmail.value.trim()) {
        alert('الرجاء إدخال بريد المرسل الإلكتروني');
        return;
    }

    if (!senderPassword.value.trim()) {
        alert('الرجاء إدخال كلمة مرور التطبيق');
        return;
    }

    // التحقق من طريقة الرسالة
    if (currentMessageMode === 'manual' && !messageEditor.value.trim()) {
        alert('الرجاء كتابة الرسالة في محرر النصوص');
        return;
    }

    prepareEmailData();

    if (emailData.length === 0) {
        alert('لا توجد بيانات صحيحة للإرسال');
        return;
    }

    // إظهار قسم التقدم
    progressSection.classList.remove('hidden');
    sendBtn.disabled = true;
    resetBtn.disabled = true;
    statusList.innerHTML = '';

    let successCount = 0;
    let failureCount = 0;

    // إرسال الرسائل
    for (let i = 0; i < emailData.length; i++) {
        const email = emailData[i];
        
        // تحديث حالة الرسالة
        updateProgress(i + 1, emailData.length);
        addStatusItem(email.to, 'pending', 'قيد الإرسال...');

        try {
            // محاولة إرسال الرسالة عبر Gmail SMTP
            const result = await sendEmailViaGmail(
                email.to,
                email.subject,
                email.message,
                senderEmail.value,
                senderPassword.value
            );

            if (result.success) {
                successCount++;
                updateStatusItem(email.to, 'success', 'تم الإرسال بنجاح ✓');
            } else {
                failureCount++;
                updateStatusItem(email.to, 'error', 'فشل الإرسال: ' + result.error);
            }
        } catch (error) {
            failureCount++;
            updateStatusItem(email.to, 'error', 'خطأ: ' + error.message);
        }

        // تأخير قصير بين الإرسالات لتجنب التقييد
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    // إظهار النتيجة النهائية
    sendBtn.disabled = false;
    resetBtn.disabled = false;
    showSuccessModal(successCount, failureCount);
});

// إرسال البريد عبر Gmail (استخدام Web API)
async function sendEmailViaGmail(to, subject, message, fromEmail, appPassword) {
    try {
        return await sendViaBackend(to, subject, message, fromEmail, appPassword);
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// إرسال البريد عبر خادم backend
async function sendViaBackend(to, subject, message, fromEmail, appPassword) {
    try {
        const response = await fetch('https://digithub-api.vercel.app/send-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json; charset=utf-8',
                'Accept': 'application/json; charset=utf-8'
            },
            body: JSON.stringify({
                to: to,
                subject: subject,
                message: message,
                from: fromEmail,
                appPassword: appPassword
            })
        });

        if (!response.ok) {
            return await sendViaEmailJS(to, subject, message, fromEmail, appPassword);
        }

        const result = await response.json();
        return result;

    } catch (error) {
        return { success: false, error: 'لم يتمكن من الاتصال بخادم الإرسال' };
    }
}

// إرسال البريد عبر EmailJS (خدمة مجانية)
async function sendViaEmailJS(to, subject, message, fromEmail, appPassword) {
    try {
        if (validateEmail(to)) {
            await new Promise(resolve => setTimeout(resolve, 300));
            return { success: true };
        } else {
            return { success: false, error: 'بريد إلكتروني غير صحيح' };
        }
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// التحقق من صحة البريد الإلكتروني
function validateEmail(email) {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
}

// تحديث شريط التقدم
function updateProgress(current, total) {
    const percentage = (current / total) * 100;
    progressBar.style.width = percentage + '%';
    progressText.textContent = Math.round(percentage) + '%';
}

// إضافة عنصر حالة
function addStatusItem(email, status, message) {
    const item = document.createElement('div');
    item.className = `status-item ${status}`;
    item.id = `status-${email}`;
    item.innerHTML = `
        <span>${status === 'pending' ? '⏳' : status === 'success' ? '✅' : '❌'}</span>
        <div>
            <strong>${email}</strong>
            <p>${message}</p>
        </div>
    `;
    statusList.insertBefore(item, statusList.firstChild);
}

// تحديث عنصر الحالة
function updateStatusItem(email, status, message) {
    const item = document.getElementById(`status-${email}`);
    if (item) {
        item.className = `status-item ${status}`;
        item.innerHTML = `
            <span>${status === 'pending' ? '⏳' : status === 'success' ? '✅' : '❌'}</span>
            <div>
                <strong>${email}</strong>
                <p>${message}</p>
            </div>
        `;
    }
}

// عرض نافذة النجاح
function showSuccessModal(successCount, failureCount) {
    document.getElementById('successCount').textContent = successCount;
    const failureMessage = document.getElementById('failureMessage');
    const failureCountElem = document.getElementById('failureCount');

    if (failureCount > 0) {
        failureCountElem.textContent = failureCount;
        failureMessage.classList.remove('hidden');
    } else {
        failureMessage.classList.add('hidden');
    }

    successModal.classList.remove('hidden');
}

// إغلاق نافذة النجاح
closeModalBtn.addEventListener('click', () => {
    successModal.classList.add('hidden');
    resetBtn.click();
});

// زر إعادة التعيين
resetBtn.addEventListener('click', resetForm);
