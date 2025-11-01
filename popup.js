document.addEventListener('DOMContentLoaded', async function() {
    let selectedDates = new Map(); // 改為 Map，儲存 {date: {type: 'both'|'in'|'out'}}
    let currentMonth, currentYear;
    let settings = {};
    let attendanceData = {}; // 儲存出勤資料

    const monthInfo = document.getElementById('monthInfo');
    const calendarGrid = document.getElementById('calendarGrid');
    const selectedInfo = document.getElementById('selectedInfo');
    const submitBtn = document.getElementById('submitBtn');
    const settingsLink = document.getElementById('settingsLink');
    const statusDiv = document.getElementById('status');
    const prevMonthBtn = document.getElementById('prevMonthBtn');
    const nextMonthBtn = document.getElementById('nextMonthBtn');
    const debugMessages = document.getElementById('debugMessages');
    const errorOverlay = document.getElementById('errorOverlay');
    
    // Debug 控制開關
    const DEBUG_MODE = false;
    const loadingOverlay = document.getElementById('loadingOverlay');
    const contextMenu = document.getElementById('contextMenu');
    let contextTarget = null; // 右鍵點擊的目標日期

    // Debug 函數，可透過 DEBUG_MODE 控制
    function addDebugMessage(message) {
        if (!DEBUG_MODE) return;
        const timestamp = new Date().toLocaleTimeString();
        debugMessages.innerHTML += `<div>[${timestamp}] ${message}</div>`;
        debugMessages.scrollTop = debugMessages.scrollHeight;
    }

    // 右鍵選單事件處理
    contextMenu.addEventListener('click', function(e) {
        const action = e.target.dataset.action;
        if (!action || !contextTarget) return;

        const dateStr = contextTarget.dataset.date;
        
        if (action === 'remove') {
            selectedDates.delete(dateStr);
        } else {
            selectedDates.set(dateStr, { type: action });
        }
        
        hideContextMenu();
        updateCalendarDisplay();
        updateSelectedInfo();
    });

    // 點擊其他地方隱藏右鍵選單
    document.addEventListener('click', function(e) {
        if (!contextMenu.contains(e.target)) {
            hideContextMenu();
        }
    });

    // 防止右鍵選單被瀏覽器右鍵選單蓋過
    document.addEventListener('contextmenu', function(e) {
        if (contextMenu.style.display === 'block') {
            e.preventDefault();
        }
    });

    function showContextMenu(e, target) {
        contextTarget = target;
        const dateKey = target.dataset.date;
        const day = parseInt(dateKey.split('-')[2]);
        
        // 檢查可用的打卡選項
        const today = new Date();
        const currentDate = new Date(currentYear, currentMonth, day);
        const isToday = currentDate.toDateString() === today.toDateString();
        
        let availableOptions = [];
        
        if (isToday) {
            // 當天：根據時間限制
            const now = new Date();
            const randomTimes = generateRandomPunchTimes(dateKey);
            
            const [startHour, startMinute] = randomTimes.startTime.split(':');
            const [endHour, endMinute] = randomTimes.endTime.split(':');
            
            const startTime = new Date();
            startTime.setHours(parseInt(startHour), parseInt(startMinute), 0, 0);
            
            const endTime = new Date();
            endTime.setHours(parseInt(endHour), parseInt(endMinute), 0, 0);
            
            const canPunchIn = startTime <= now;
            const canPunchOut = endTime <= now;
            
            if (canPunchIn && canPunchOut) {
                availableOptions = ['both', 'in', 'out'];
            } else if (canPunchIn) {
                availableOptions = ['in'];
            } else {
                availableOptions = [];
            }
        } else if (currentDate < today) {
            // 過去日期：全部選項都可用
            availableOptions = ['both', 'in', 'out'];
        } else {
            // 未來日期：不可用
            availableOptions = [];
        }
        
        // 動態顯示/隱藏選單項目
        const menuItems = contextMenu.querySelectorAll('.context-menu-item[data-action]');
        menuItems.forEach(item => {
            const action = item.dataset.action;
            if (action === 'remove' || availableOptions.includes(action)) {
                item.style.display = 'block';
            } else {
                item.style.display = 'none';
            }
        });
        
        if (availableOptions.length === 0) {
            // 如果沒有可用選項，不顯示選單
            return;
        }
        
        contextMenu.style.display = 'block';
        contextMenu.style.left = e.pageX + 'px';
        contextMenu.style.top = e.pageY + 'px';
        
        // 確保選單不會超出視窗
        const rect = contextMenu.getBoundingClientRect();
        if (rect.right > window.innerWidth) {
            contextMenu.style.left = (e.pageX - rect.width) + 'px';
        }
        if (rect.bottom > window.innerHeight) {
            contextMenu.style.top = (e.pageY - rect.height) + 'px';
        }
    }

    function hideContextMenu() {
        contextMenu.style.display = 'none';
        contextTarget = null;
    }

    // 控制 debug 區域顯示
    const debugArea = document.getElementById('debugArea');
    if (DEBUG_MODE) {
        debugArea.style.display = 'block';
    }

    // 載入設定並初始化
    await loadSettings();
    initializeCalendar();

    // 事件監聽器
    settingsLink.addEventListener('click', function(e) {
        e.preventDefault();
        chrome.runtime.openOptionsPage();
    });

    submitBtn.addEventListener('click', handleSubmit);
    
    
    prevMonthBtn.addEventListener('click', async function() {
        prevMonthBtn.disabled = true;
        nextMonthBtn.disabled = true;
        showLoading();
        
        try {
            currentMonth--;
            if (currentMonth < 0) {
                currentMonth = 11;
                currentYear--;
            }
            selectedDates.clear(); // 清空已選日期
            await loadAttendanceData(); // 載入出勤資料
            renderCalendar();
        } finally {
            hideLoading();
            prevMonthBtn.disabled = false;
            nextMonthBtn.disabled = false;
        }
    });
    
    nextMonthBtn.addEventListener('click', async function() {
        prevMonthBtn.disabled = true;
        nextMonthBtn.disabled = true;
        showLoading();
        
        try {
            currentMonth++;
            if (currentMonth > 11) {
                currentMonth = 0;
                currentYear++;
            }
            selectedDates.clear(); // 清空已選日期
            await loadAttendanceData(); // 載入出勤資料
            renderCalendar();
        } finally {
            hideLoading();
            prevMonthBtn.disabled = false;
            nextMonthBtn.disabled = false;
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
            settings = result;

        } catch (error) {
            showStatus('載入設定失敗: ' + error.message, 'error');
        }
    }

    function initializeCalendar() {
        const now = new Date();

        // 打開當日所在月份（使用本地時區）
        currentYear = now.getFullYear();
        currentMonth = now.getMonth();

        loadAttendanceDataAndRender();
    }

    async function loadAttendanceDataAndRender() {
        // TODO: 除錯用，記得之後移除
        addDebugMessage('開始載入出勤資料...');
        await loadAttendanceData();
        addDebugMessage('資料載入完成，開始渲染月曆...');
        renderCalendar();
        addDebugMessage('月曆渲染完成');
    }

    async function loadAttendanceData() {
        if (!settings.employeeId) {
            // TODO: 除錯用，記得之後移除
            addDebugMessage('沒有員工 ID，跳過載入');
            return; // 沒有員工 ID 就不載入
        }
        
        // TODO: 除錯用，記得之後移除
        addDebugMessage(`找到員工 ID: ${settings.employeeId}，準備載入資料...`);

        try {
            // 計算月份的第一天和最後一天
            const firstDay = new Date(currentYear, currentMonth, 1);
            const lastDay = new Date(currentYear, currentMonth + 1, 0);
            
            // 避免時區問題，手動格式化日期
            const startDate = `${firstDay.getFullYear()}-${String(firstDay.getMonth() + 1).padStart(2, '0')}-${String(firstDay.getDate()).padStart(2, '0')}`;
            const endDate = `${lastDay.getFullYear()}-${String(lastDay.getMonth() + 1).padStart(2, '0')}-${String(lastDay.getDate()).padStart(2, '0')}`;
            
            // TODO: 除錯用，記得之後移除
            addDebugMessage(`載入出勤資料 - 日期範圍: ${startDate} 到 ${endDate}`);
            addDebugMessage(`當前年月: ${currentYear}-${currentMonth + 1}，月份最後一天: ${lastDay.getDate()}`);
            
            // 獲取當前頁面
            const [tab] = await chrome.tabs.query({ url: '*://cloud.nueip.com/*' });
            
            if (tab) {
                const response = await chrome.tabs.sendMessage(tab.id, {
                    action: 'fetchAttendanceData',
                    data: {
                        startDate: startDate,
                        endDate: endDate,
                        employeeId: settings.employeeId
                    }
                });
                
                if (response && response.success) {
                    attendanceData = response.data;
                    
                    // TODO: 除錯用，記得之後移除
                    const dates = Object.keys(attendanceData).sort();
                    addDebugMessage(`載入成功！總共 ${dates.length} 天資料`);
                    addDebugMessage(`日期清單: ${dates.join(', ')}`);
                } else {
                    addDebugMessage(`載入失敗: ${response?.error || '未知錯誤'}`);
                    showLoadingError();
                }
            } else {
                addDebugMessage('找不到人資系統頁面');
                showLoadingError();
            }
        } catch (error) {
            addDebugMessage(`載入錯誤: ${error.message}`);
            showLoadingError();
        }
    }

    function showLoadingError() {
        errorOverlay.style.display = 'flex';
    }

    function showLoading() {
        loadingOverlay.style.display = 'flex';
    }

    function hideLoading() {
        loadingOverlay.style.display = 'none';
    }

    function renderCalendar() {
        const monthNames = [
            '一月', '二月', '三月', '四月', '五月', '六月',
            '七月', '八月', '九月', '十月', '十一月', '十二月'
        ];

        monthInfo.textContent = `${currentYear} 年 ${monthNames[currentMonth]}`;

        // 清空日曆
        calendarGrid.innerHTML = '';

        // 計算該月的第一天和最後一天
        const firstDay = new Date(currentYear, currentMonth, 1);
        const lastDay = new Date(currentYear, currentMonth + 1, 0);
        const firstDayOfWeek = firstDay.getDay();
        const daysInMonth = lastDay.getDate();

        // 計算上個月需要顯示的天數
        const prevMonth = new Date(currentYear, currentMonth, 0);
        const daysInPrevMonth = prevMonth.getDate();

        // 添加上個月的日期
        for (let i = firstDayOfWeek - 1; i >= 0; i--) {
            const day = daysInPrevMonth - i;
            let prevYear = currentYear;
            let prevMonthIndex = currentMonth - 1;
            if (prevMonthIndex < 0) {
                prevMonthIndex = 11;
                prevYear = currentYear - 1;
            }
            const dayElement = createDayElement(day, true, null, prevYear, prevMonthIndex);
            calendarGrid.appendChild(dayElement);
        }

        // 添加本月的日期
        for (let day = 1; day <= daysInMonth; day++) {
            const dayOfWeek = (firstDayOfWeek + day - 1) % 7;
            const dayElement = createDayElement(day, false, dayOfWeek);
            calendarGrid.appendChild(dayElement);
        }

        // 計算剩餘格子並添加下個月的日期
        const totalCells = calendarGrid.children.length;
        const remainingCells = (42 - totalCells) % 7;
        if (remainingCells > 0) {
            for (let day = 1; day <= (7 - remainingCells); day++) {
                let nextYear = currentYear;
                let nextMonthIndex = currentMonth + 1;
                if (nextMonthIndex > 11) {
                    nextMonthIndex = 0;
                    nextYear = currentYear + 1;
                }
                const dayElement = createDayElement(day, true, null, nextYear, nextMonthIndex);
                calendarGrid.appendChild(dayElement);
            }
        }

        updateSelectedInfo();
    }

    function createDayElement(day, isOtherMonth, dayOfWeek = null, actualYear = null, actualMonth = null) {
        const dayElement = document.createElement('div');
        dayElement.className = 'calendar-day';
        dayElement.textContent = day;
        
        // 設定正確的日期鍵（針對其他月份使用正確的年月）
        const year = actualYear || currentYear;
        const month = actualMonth !== null ? actualMonth : currentMonth;
        const dateKey = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        dayElement.dataset.date = dateKey;

        if (isOtherMonth) {
            dayElement.classList.add('other-month');
            return dayElement;
        }
        
        // 解析出勤資料狀態
        const dayStatus = getDayStatus(dateKey, dayOfWeek);
        
        // 添加狀態樣式
        if (dayStatus.isHoliday) {
            dayElement.classList.add('holiday');
        } else if (dayStatus.isWeekend) {
            dayElement.classList.add('weekend');
        } else {
            dayElement.classList.add('workday');
        }
        
        if (dayStatus.hasTimeoff) {
            dayElement.classList.add('timeoff');
        }
        
        // 處理打卡狀態的分割顯示
        if (dayStatus.hasPunchIn && dayStatus.hasPunchOut) {
            dayElement.classList.add('punched-both');
        } else if (dayStatus.hasPunchIn) {
            dayElement.classList.add('punched-in');
        } else if (dayStatus.hasPunchOut) {
            dayElement.classList.add('punched-out');
        }

        
        // 只在有複雜狀態時才設定 tooltip
        if (dayStatus.hasTimeoff || dayStatus.hasPunchIn || dayStatus.hasPunchOut || dayStatus.isHoliday) {
            dayElement.title = dayStatus.generateTooltip();
        }

        // 檢查是否已選擇
        const selection = selectedDates.get(dateKey);
        if (selection) {
            dayElement.classList.add('selected');
            dayElement.classList.add('selected-' + selection.type);
        }

        // 添加右鍵事件
        dayElement.addEventListener('contextmenu', function(e) {
            e.preventDefault();
            
            // 檢查是否已設定員工 ID
            if (!settings.employeeId) {
                showStatus('請先到設定頁面設定員工 ID', 'error');
                return;
            }
            
            const isCurrentMonth = !isOtherMonth;
            const dayType = dayStatus.isWeekend ? 'weekend' : dayStatus.isHoliday ? 'holiday' : 'workday';
            if (!isDaySelectable(day, isCurrentMonth, dayType)) return;
            showContextMenu(e, dayElement);
        });
        
        // 保留左鍵點擊用於智能選擇
        dayElement.addEventListener('click', function() {
            // 檢查是否已設定員工 ID
            if (!settings.employeeId) {
                showStatus('請先到設定頁面設定員工 ID', 'error');
                return;
            }
            
            const isCurrentMonth = !isOtherMonth;
            const dayType = dayStatus.isWeekend ? 'weekend' : dayStatus.isHoliday ? 'holiday' : 'workday';
            if (!isDaySelectable(day, isCurrentMonth, dayType)) return;
            smartLeftClickSelection(dateKey, dayElement, day);
        });

        return dayElement;
    }

    function generateRandomPunchTimes(dateKey) {
        // 檢查是否已經有打卡記錄，如果有就使用相同的分鐘數
        const dayData = attendanceData[dateKey]?.[settings.employeeId];
        let existingMinute = null;
        
        if (dayData && dayData.punch) {
            // 檢查是否已有上班卡
            if (dayData.punch.onPunch && dayData.punch.onPunch.length > 0) {
                const existingTime = dayData.punch.onPunch[0].time;
                if (existingTime && existingTime.includes(':')) {
                    existingMinute = existingTime.split(':')[1];
                    addDebugMessage(`${dateKey}: 使用現有上班卡分鐘數 ${existingMinute}`);
                }
            }
            // 檢查是否已有下班卡
            else if (dayData.punch.offPunch && dayData.punch.offPunch.length > 0) {
                const existingTime = dayData.punch.offPunch[0].time;
                if (existingTime && existingTime.includes(':')) {
                    existingMinute = existingTime.split(':')[1];
                    addDebugMessage(`${dateKey}: 使用現有下班卡分鐘數 ${existingMinute}`);
                }
            }
        }
        
        // 如果沒有現有記錄，生成新的隨機分鐘數
        const minute = existingMinute || String(Math.floor(Math.random() * 59)).padStart(2, '0');
        
        if (!existingMinute) {
            addDebugMessage(`${dateKey}: 生成新的隨機分鐘數 ${minute}`);
        }
        
        return {
            startTime: `${settings.startHour}:${minute}`,
            endTime: `${settings.endHour}:${minute}`
        };
    }

    function isDaySelectable(day, isCurrentMonth, dayType, checkFutureTime = false, punchType = 'both') {
        if (!isCurrentMonth || dayType !== 'workday') {
            return false;
        }
        
        const today = new Date();
        const currentDate = new Date(currentYear, currentMonth, day);
        
        // 檢查未來日期
        if (currentDate > today) {
            if (checkFutureTime) {
                showStatus('不能選擇未來日期', 'error');
            }
            return false;
        }
        
        // 檢查未來時間（同一天但時間超過現在）
        if (checkFutureTime && currentDate.toDateString() === today.toDateString()) {
            const now = new Date();
            const dateKey = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            
            // 生成實際會使用的隨機時間
            const randomTimes = generateRandomPunchTimes(dateKey);
            
            // 檢查上班時間是否為未來
            if (punchType === 'both' || punchType === 'in') {
                const [hour, minute] = randomTimes.startTime.split(':');
                const startTime = new Date();
                startTime.setHours(parseInt(hour), parseInt(minute), 0, 0);
                
                if (startTime > now) {
                    showStatus('上班打卡時間還未到', 'error');
                    return false;
                }
            }
            
            // 檢查下班時間是否為未來
            if (punchType === 'both' || punchType === 'out') {
                const [hour, minute] = randomTimes.endTime.split(':');
                const endTime = new Date();
                endTime.setHours(parseInt(hour), parseInt(minute), 0, 0);
                
                if (endTime > now) {
                    showStatus('下班打卡時間還未到', 'error');
                    return false;
                }
            }
        }
        
        return true;
    }

    function getDayStatus(dateKey, dayOfWeek = null) {
        const status = {
            isWeekend: false,
            isHoliday: false,
            holidayName: '',
            hasTimeoff: false,
            timeoffDetails: [],
            hasPunchIn: false,
            hasPunchOut: false,
            punchInTimes: [],
            punchOutTimes: [],
            tooltip: ''
        };

        // 檢查是否為週末（使用傳入的 dayOfWeek 避免重複創建 Date 物件）
        if (dayOfWeek !== null) {
            status.isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        } else {
            const date = new Date(dateKey);
            status.isWeekend = date.getDay() === 0 || date.getDay() === 6;
        }

        // 檢查出勤資料
        const dayData = attendanceData[dateKey]?.[settings.employeeId];
        if (dayData) {
            
            // 檢查是否為假日
            if (dayData.dateInfo.date_off) {
                status.isHoliday = true;
                if (dayData.dateInfo.holiday) {
                    status.holidayName = dayData.dateInfo.holiday;
                }
            }
            
            // 檢查是否有請假
            if (dayData.timeoff && dayData.timeoff.length > 0) {
                status.hasTimeoff = true;
                status.timeoffDetails = dayData.timeoff.map(item => ({
                    ruleName: item.rule_name,
                    time: item.time,
                    remark: item.remark
                }));
            }
            
            // 檢查打卡記錄
            if (dayData.punch) {
                if (dayData.punch.onPunch && dayData.punch.onPunch.length > 0) {
                    status.hasPunchIn = true;
                    status.punchInTimes = dayData.punch.onPunch.map(p => p.time);
                }
                if (dayData.punch.offPunch && dayData.punch.offPunch.length > 0) {
                    status.hasPunchOut = true;
                    status.punchOutTimes = dayData.punch.offPunch.map(p => p.time);
                }
            }
        }

        // 延遲生成 tooltip（只在需要時生成）
        status.generateTooltip = () => generateTooltip(status, dateKey);

        return status;
    }

    function generateTooltip(status, dateKey) {
        const parts = [];
        
        parts.push(`日期: ${dateKey}`);
        
        if (status.isHoliday && status.holidayName) {
            parts.push(`國定假日: ${status.holidayName}`);
        } else if (status.isWeekend) {
            parts.push('週末');
        } else {
            parts.push('工作日');
        }
        
        if (status.hasTimeoff && status.timeoffDetails.length > 0) {
            parts.push('請假記錄:');
            status.timeoffDetails.forEach(item => {
                parts.push(`• ${item.ruleName} (${item.time}) - ${item.remark}`);
            });
        }
        
        if (status.hasPunchIn || status.hasPunchOut) {
            parts.push('打卡記錄:');
            if (status.hasPunchIn) {
                parts.push(`• 上班: ${status.punchInTimes.join(', ')}`);
            }
            if (status.hasPunchOut) {
                parts.push(`• 下班: ${status.punchOutTimes.join(', ')}`);
            }
        }
        
        return parts.join('\n');
    }

    function toggleDateSelection(dateKey, dayElement, punchType = 'both') {
        if (selectedDates.has(dateKey)) {
            selectedDates.delete(dateKey);
        } else {
            selectedDates.set(dateKey, { type: punchType });
        }
        updateCalendarDisplay();
        updateSelectedInfo();
    }

    function smartLeftClickSelection(dateKey, dayElement, day) {
        // 如果已經選擇過這個日期，取消選擇
        if (selectedDates.has(dateKey)) {
            selectedDates.delete(dateKey);
            updateCalendarDisplay();
            updateSelectedInfo();
            return;
        }

        // 檢查當天是否已經有打卡記錄
        const dayData = attendanceData[dateKey]?.[settings.employeeId];
        let selectedType = 'both'; // 預設值
        
        if (dayData && dayData.punch) {
            const hasPunchIn = dayData.punch.onPunch && dayData.punch.onPunch.length > 0;
            const hasPunchOut = dayData.punch.offPunch && dayData.punch.offPunch.length > 0;

            if (hasPunchIn && !hasPunchOut) {
                // 已經打過上班卡，智能選擇下班卡
                selectedType = 'out';
                addDebugMessage(`${dateKey}: 已有上班卡，智能選擇下班卡`);
            } else if (!hasPunchIn && hasPunchOut) {
                // 已經打過下班卡，智能選擇上班卡
                selectedType = 'in';
                addDebugMessage(`${dateKey}: 已有下班卡，智能選擇上班卡`);
            } else if (hasPunchIn && hasPunchOut) {
                // 兩個都打過，選擇補兩張卡
                selectedType = 'both';
                addDebugMessage(`${dateKey}: 已有上下班卡，選擇補兩張卡`);
            } else {
                // 都沒打過，根據時間智能選擇
                const today = new Date();
                const currentDate = new Date(currentYear, currentMonth, day);
                
                if (currentDate.toDateString() === today.toDateString()) {
                    // 當天：根據現在時間決定
                    const now = new Date();
                    const randomTimes = generateRandomPunchTimes(dateKey);
                    
                    const [startHour, startMinute] = randomTimes.startTime.split(':');
                    const [endHour, endMinute] = randomTimes.endTime.split(':');
                    
                    const startTime = new Date();
                    startTime.setHours(parseInt(startHour), parseInt(startMinute), 0, 0);
                    
                    const endTime = new Date();
                    endTime.setHours(parseInt(endHour), parseInt(endMinute), 0, 0);
                    
                    const canPunchIn = startTime <= now;
                    const canPunchOut = endTime <= now;
                    
                    if (canPunchIn && canPunchOut) {
                        selectedType = 'both';
                        addDebugMessage(`${dateKey}: 當天，上下班時間都已過，選擇補兩張卡`);
                    } else if (canPunchIn && !canPunchOut) {
                        selectedType = 'in';
                        addDebugMessage(`${dateKey}: 當天，只有上班時間已過，選擇補上班卡`);
                    } else {
                        // 都還沒到時間，不應該能選
                        addDebugMessage(`${dateKey}: 當天，上班時間還沒到`);
                        return;
                    }
                } else {
                    // 過去日期：預設補兩張卡
                    selectedType = 'both';
                    addDebugMessage(`${dateKey}: 過去日期，選擇補上下班卡`);
                }
            }
        } else {
            // 沒有出勤資料，使用相同的時間邏輯
            const today = new Date();
            const currentDate = new Date(currentYear, currentMonth, day);
            
            if (currentDate.toDateString() === today.toDateString()) {
                const now = new Date();
                const randomTimes = generateRandomPunchTimes(dateKey);
                
                const [startHour, startMinute] = randomTimes.startTime.split(':');
                const [endHour, endMinute] = randomTimes.endTime.split(':');
                
                const startTime = new Date();
                startTime.setHours(parseInt(startHour), parseInt(startMinute), 0, 0);
                
                const endTime = new Date();
                endTime.setHours(parseInt(endHour), parseInt(endMinute), 0, 0);
                
                const canPunchIn = startTime <= now;
                const canPunchOut = endTime <= now;
                
                if (canPunchIn && canPunchOut) {
                    selectedType = 'both';
                    addDebugMessage(`${dateKey}: 無資料當天，上下班時間都已過，選擇補兩張卡`);
                } else if (canPunchIn && !canPunchOut) {
                    selectedType = 'in';
                    addDebugMessage(`${dateKey}: 無資料當天，只有上班時間已過，選擇補上班卡`);
                } else {
                    addDebugMessage(`${dateKey}: 無資料當天，上班時間還沒到`);
                    return;
                }
            } else {
                selectedType = 'both';
                addDebugMessage(`${dateKey}: 無資料過去日期，選擇補上下班卡`);
            }
        }

        selectedDates.set(dateKey, { type: selectedType });
        updateCalendarDisplay();
        updateSelectedInfo();
    }

    function updateCalendarDisplay() {
        const dayElements = calendarGrid.querySelectorAll('.calendar-day[data-date]');
        dayElements.forEach(dayElement => {
            const dateKey = dayElement.dataset.date;
            const selection = selectedDates.get(dateKey);
            
            // 清除所有選擇相關的 class
            dayElement.classList.remove('selected', 'selected-in', 'selected-out', 'selected-both');
            
            if (selection) {
                dayElement.classList.add('selected');
                dayElement.classList.add('selected-' + selection.type);
            }
        });
    }

    function updateSelectedInfo() {
        // 如果沒有員工 ID，顯示錯誤訊息
        if (!settings.employeeId) {
            selectedInfo.innerHTML = `
                <div style="color: #dc3545; font-weight: 500;">
                    尚未設定員工 ID
                </div>
                <small>請到設定頁面完成設定</small>
            `;
            submitBtn.disabled = true;
            submitBtn.textContent = '請先完成設定';
            return;
        }

        const count = selectedDates.size;
        if (count === 0) {
            selectedInfo.textContent = '點選日期來選擇需要補卡的日期（左鍵：上下班卡，右鍵：更多選項）';
            submitBtn.disabled = true;
            submitBtn.textContent = '送出補卡申請';
        } else {
            let totalPunches = 0;
            let details = [];
            
            selectedDates.forEach((selection, date) => {
                const punchCount = selection.type === 'both' ? 2 : 1;
                totalPunches += punchCount;
                
                const typeText = selection.type === 'both' ? '上下班卡' : 
                               selection.type === 'in' ? '上班卡' : '下班卡';
                details.push(`${date}: ${typeText}`);
            });
            
            selectedInfo.innerHTML = `已選擇 ${count} 天，共 ${totalPunches} 張卡<br><small>${details.join('<br>')}</small>`;
            submitBtn.disabled = false;
            submitBtn.textContent = '送出補卡申請';
        }
    }

    async function handleSubmit() {
        if (selectedDates.size === 0) {
            showStatus('請選擇至少一個日期', 'error');
            return;
        }

        if (!settings.employeeId) {
            showStatus('請先到設定頁面設定員工 ID', 'error');
            return;
        }

        submitBtn.disabled = true;
        submitBtn.textContent = '送出中...';
        
        try {
            // 獲取當前頁面
            const tabs = await chrome.tabs.query({ url: '*://cloud.nueip.com/*' });
            const tab = tabs[0];
            
            // 準備補卡資料
            const punchData = [];
            
            selectedDates.forEach((selection, dateKey) => {
                const [year, month, day] = dateKey.split('-');
                const dayNum = parseInt(day);
                
                // 使用相同的時間生成邏輯（保持與現有打卡記錄一致）
                const times = generateRandomPunchTimes(dateKey);
                
                const punchItem = {
                    date: dateKey,
                    day: dayNum,
                    punchType: selection.type
                };
                
                if (selection.type === 'both' || selection.type === 'in') {
                    punchItem.startTime = times.startTime;
                }
                if (selection.type === 'both' || selection.type === 'out') {
                    punchItem.endTime = times.endTime;
                }
                
                punchData.push(punchItem);
            });

            // 發送到 content script
            const response = await chrome.tabs.sendMessage(tab.id, {
                action: 'submitPunchCards',
                data: {
                    punchData: punchData,
                    settings: settings
                }
            });

            if (response.success) {
                showStatus(`成功送出 ${response.total} 筆補卡申請`, 'success');
                // 清空選擇
                selectedDates.clear();
                
                // 重新載入出勤資料並重新渲染月曆
                await loadAttendanceData();
                renderCalendar();
                updateSelectedInfo();
            } else {
                throw new Error(response.error || '送出失敗');
            }

        } catch (error) {
            showStatus('送出失敗: ' + error.message, 'error');
        } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = '送出補卡申請';
        }
    }

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = `status ${type}`;
        statusDiv.style.display = 'block';
        
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
});