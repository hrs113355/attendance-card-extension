document.addEventListener('DOMContentLoaded', async function() {
    const employeeIdInput = document.getElementById('employeeId');
    const reasonInput = document.getElementById('reason');
    const startHourSelect = document.getElementById('startHour');
    const endHourSelect = document.getElementById('endHour');
    const saveBtn = document.getElementById('saveBtn');
    const statusDiv = document.getElementById('status');
    const fetchIdBtn = document.getElementById('fetchIdBtn');

    // 載入已儲存的設定
    await loadSettings();


    // 自動抓取員工 ID
    fetchIdBtn.addEventListener('click', async function() {
        fetchIdBtn.disabled = true;
        fetchIdBtn.textContent = '抓取中...';
        
        try {
            // 查詢所有 nueip 相關的 tab
            const tabs = await chrome.tabs.query({ url: '*://cloud.nueip.com/*' });
            console.log('找到的 tabs:', tabs);
            
            if (tabs.length === 0) {
                throw new Error('請先登入人資系統頁面：<a href="https://cloud.nueip.com" target="_blank">https://cloud.nueip.com</a>');
            }
            
            // 使用第一個找到的 nueip tab
            const tab = tabs[0];
            console.log('使用的 tab:', tab);
            
            // 發送訊息到 content script
            console.log('準備發送訊息到 tab:', tab.id, tab.url);
            const response = await chrome.tabs.sendMessage(tab.id, {
                action: 'fetchEmployeeId'
            }).catch(error => {
                console.error('sendMessage 錯誤:', error);
                throw new Error('Content script 未回應，請重新整理頁面後再試');
            });
            
            console.log('收到回應:', response);
            
            if (response && response.success) {
                employeeIdInput.value = response.employeeId;
                showStatus('員工 ID 自動抓取成功: ' + response.employeeId, 'success');
            } else {
                throw new Error(response?.error || 'Content script 沒有回應或抓取失敗');
            }
            
        } catch (error) {
            // 如果錯誤訊息包含 HTML，使用 allowHTML 參數
            const hasHTML = error.message.includes('<a href');
            showStatus('抓取失敗: ' + error.message, 'error', hasHTML);
        } finally {
            fetchIdBtn.disabled = false;
            fetchIdBtn.textContent = '自動抓取';
        }
    });

    // 儲存設定
    saveBtn.addEventListener('click', async function() {
        const settings = {
            employeeId: employeeIdInput.value.trim(),
            reason: reasonInput.value.trim() || '補卡',
            startHour: startHourSelect.value,
            endHour: endHourSelect.value,
            autoMonth: true
        };

        // 驗證必填欄位
        if (!settings.employeeId) {
            showStatus('請輸入員工 ID', 'error');
            return;
        }

        try {
            await chrome.storage.sync.set(settings);
            showStatus('設定已儲存', 'success');
        } catch (error) {
            showStatus('儲存失敗: ' + error.message, 'error');
        }
    });

    async function loadSettings() {
        try {
            const result = await chrome.storage.sync.get({
                employeeId: '',
                reason: '補卡',
                startHour: '09',
                endHour: '18',
                autoMonth: true
            });

            employeeIdInput.value = result.employeeId;
            reasonInput.value = result.reason;
            startHourSelect.value = result.startHour;
            endHourSelect.value = result.endHour;
        } catch (error) {
            showStatus('載入設定失敗: ' + error.message, 'error');
        }
    }

    function showStatus(message, type, allowHTML = false) {
        if (allowHTML) {
            statusDiv.innerHTML = message;
        } else {
            statusDiv.textContent = message;
        }
        statusDiv.className = `status ${type}`;
        statusDiv.style.display = 'block';
        
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 3000);
    }
});