// Content script for handling punch card submissions
(function() {
    'use strict';
    
    console.log('補卡助手 Content Script 已載入在:', window.location.href);
    console.log('Document ready state:', document.readyState);

    // 測試訊息監聽器是否設置成功
    console.log('設置訊息監聽器...');
    
    // 監聽來自 popup 和 options 的訊息
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        console.log('收到訊息:', request);
        if (request.action === 'submitPunchCards') {
            submitPunchCards(request.data)
                .then(result => sendResponse(result))
                .catch(error => sendResponse({ success: false, error: error.message }));
            
            // 返回 true 表示會非同步回應
            return true;
        } else if (request.action === 'fetchEmployeeId') {
            console.log('收到 fetchEmployeeId 請求');
            fetchEmployeeId()
                .then(result => {
                    console.log('fetchEmployeeId 結果:', result);
                    sendResponse(result);
                })
                .catch(error => {
                    console.error('fetchEmployeeId 錯誤:', error);
                    sendResponse({ success: false, error: error.message });
                });
            
            // 返回 true 表示會非同步回應
            return true;
        } else if (request.action === 'fetchAttendanceData') {
            console.log('收到 fetchAttendanceData 請求:', request.data);
            fetchAttendanceData(request.data)
                .then(result => {
                    console.log('fetchAttendanceData 結果:', result);
                    sendResponse(result);
                })
                .catch(error => {
                    console.error('fetchAttendanceData 錯誤:', error);
                    sendResponse({ success: false, error: error.message });
                });
            
            // 返回 true 表示會非同步回應
            return true;
        }
    });

    async function submitPunchCards(data) {
        const { punchData, settings } = data;
        const results = [];
        let successCount = 0;
        let failCount = 0;

        let totalPunches = 0;
        punchData.forEach(punch => {
            if (punch.punchType === 'both') totalPunches += 2;
            else totalPunches += 1;
        });
        
        console.log('開始補卡作業，總共', totalPunches, '筆');

        for (const punch of punchData) {
            try {
                // 補上班卡
                if (punch.punchType === 'both' || punch.punchType === 'in') {
                    const startResult = await submitSinglePunch({
                        date: punch.date,
                        time: punch.startTime,
                        section: 1, // 上班
                        day: punch.day,
                        settings: settings
                    });

                    if (startResult.success) {
                        successCount++;
                        console.log(`${punch.date} 上班卡補卡成功 (${punch.startTime})`);
                    } else {
                        failCount++;
                        console.error(`${punch.date} 上班卡補卡失敗:`, startResult.error);
                    }

                    results.push({
                        date: punch.date,
                        type: '上班',
                        time: punch.startTime,
                        success: startResult.success,
                        error: startResult.error
                    });

                    // 延遲一下避免請求過於頻繁
                    await sleep(200);
                }

                // 補下班卡
                if (punch.punchType === 'both' || punch.punchType === 'out') {
                    const endResult = await submitSinglePunch({
                        date: punch.date,
                        time: punch.endTime,
                        section: 2, // 下班
                        day: punch.day,
                        settings: settings
                    });

                    if (endResult.success) {
                        successCount++;
                        console.log(`${punch.date} 下班卡補卡成功 (${punch.endTime})`);
                    } else {
                        failCount++;
                        console.error(`${punch.date} 下班卡補卡失敗:`, endResult.error);
                    }

                    results.push({
                        date: punch.date,
                        type: '下班',
                        time: punch.endTime,
                        success: endResult.success,
                        error: endResult.error
                    });

                    // 延遲一下避免請求過於頻繁
                    await sleep(200);
                }

            } catch (error) {
                console.error(`${punch.date} 補卡過程發生錯誤:`, error);
                const punchCount = punch.punchType === 'both' ? 2 : 1;
                failCount += punchCount;
                results.push({
                    date: punch.date,
                    type: punch.punchType === 'both' ? '上班+下班' : 
                          punch.punchType === 'in' ? '上班' : '下班',
                    success: false,
                    error: error.message
                });
            }
        }

        const summary = {
            success: failCount === 0,
            successCount: successCount,
            failCount: failCount,
            total: totalPunches,
            results: results
        };

        console.log('補卡作業完成:', summary);
        return summary;
    }

    function submitSinglePunch({ date, time, section, day, settings }) {
        return new Promise((resolve) => {
            const [hour, minute] = time.split(':');
            
            const params = {
                section: section,
                hour: hour,
                min: minute,
                remark: settings.reason,
                u_sn: settings.employeeId,
                date: date,
                apply_date: date
            };

            // 將參數編碼為 URL encoded 格式
            const urlEncodedDataPairs = [];
            for (let name in params) {
                urlEncodedDataPairs.push(
                    encodeURIComponent(name) + '=' + encodeURIComponent(params[name])
                );
            }
            const urlEncodedData = urlEncodedDataPairs.join('&');

            console.log('發送補卡請求:', params);

            const xhr = new XMLHttpRequest();
            const url = '/attendance_record/addCorrectionPunch';

            xhr.open('POST', url, true);
            xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === XMLHttpRequest.DONE) {
                    if (xhr.status === 200) {
                        try {
                            // 嘗試解析回應
                            let response = xhr.responseText;
                            
                            // 如果是 JSON 回應，嘗試解析
                            if (response.trim().startsWith('{') || response.trim().startsWith('[')) {
                                const jsonResponse = JSON.parse(response);
                                resolve({
                                    success: true,
                                    response: jsonResponse
                                });
                            } else {
                                // 如果不是 JSON，檢查是否包含成功指示
                                const isSuccess = !response.toLowerCase().includes('error') && 
                                                !response.toLowerCase().includes('fail') &&
                                                xhr.status === 200;
                                resolve({
                                    success: isSuccess,
                                    response: response
                                });
                            }
                        } catch (error) {
                            // 解析失敗但狀態碼是 200，可能仍然成功
                            resolve({
                                success: true,
                                response: xhr.responseText,
                                note: 'Response parsing failed but request succeeded'
                            });
                        }
                    } else {
                        resolve({
                            success: false,
                            error: `HTTP ${xhr.status}: ${xhr.statusText}`,
                            response: xhr.responseText
                        });
                    }
                }
            };

            xhr.onerror = function() {
                resolve({
                    success: false,
                    error: '網路錯誤或請求被阻擋'
                });
            };

            xhr.ontimeout = function() {
                resolve({
                    success: false,
                    error: '請求超時'
                });
            };

            // 設定超時時間
            xhr.timeout = 10000; // 10 秒超時

            xhr.send(urlEncodedData);
        });
    }

    async function fetchEmployeeId() {
        return new Promise((resolve) => {
            const xhr = new XMLHttpRequest();
            const url = '/attendance_record/ajax';
            
            // 構建請求參數
            const params = new URLSearchParams({
                action: 'attendance',
                loadInBatch: '1',
                loadBatchGroupNum: '1000',
                loadBatchNumber: '1',
                work_status: '1'
            });
            
            xhr.open('POST', url, true);
            xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === XMLHttpRequest.DONE) {
                    if (xhr.status === 200) {
                        try {
                            const response = JSON.parse(xhr.responseText);
                            
                            // 檢查回應格式
                            if (response && response.data) {
                                // 取得第一個日期的資料
                                const firstDate = Object.keys(response.data)[0];
                                if (firstDate && response.data[firstDate]) {
                                    // 取得第一個員工 ID
                                    const employeeId = Object.keys(response.data[firstDate])[0];
                                    if (employeeId) {
                                        resolve({
                                            success: true,
                                            employeeId: employeeId
                                        });
                                        return;
                                    }
                                }
                            }
                            
                            resolve({
                                success: false,
                                error: '無法從回應中找到員工 ID'
                            });
                            
                        } catch (error) {
                            resolve({
                                success: false,
                                error: '解析回應失敗: ' + error.message
                            });
                        }
                    } else {
                        resolve({
                            success: false,
                            error: `HTTP ${xhr.status}: ${xhr.statusText}`
                        });
                    }
                }
            };
            
            xhr.onerror = function() {
                resolve({
                    success: false,
                    error: '網路錯誤或請求被阻擋'
                });
            };
            
            xhr.ontimeout = function() {
                resolve({
                    success: false,
                    error: '請求超時'
                });
            };
            
            // 設定超時時間
            xhr.timeout = 10000; // 10 秒超時
            
            xhr.send(params.toString());
        });
    }

    async function fetchAttendanceData({ startDate, endDate, employeeId }) {
        return new Promise((resolve) => {
            const xhr = new XMLHttpRequest();
            const url = '/attendance_record/ajax';
            
            // 設定日期範圍的 cookie
            document.cookie = `Search_124_date_start=${startDate}; path=/`;
            document.cookie = `Search_124_date_end=${endDate}; path=/`;
            
            // 構建請求參數
            const params = new URLSearchParams({
                action: 'attendance',
                loadInBatch: '1',
                loadBatchGroupNum: '1000',
                loadBatchNumber: '1',
                work_status: '1'
            });
            
            xhr.open('POST', url, true);
            xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
            
            xhr.onreadystatechange = function() {
                if (xhr.readyState === XMLHttpRequest.DONE) {
                    if (xhr.status === 200) {
                        try {
                            const response = JSON.parse(xhr.responseText);
                            
                            if (response && response.data) {
                                resolve({
                                    success: true,
                                    data: response.data,
                                    employeeId: employeeId
                                });
                            } else {
                                resolve({
                                    success: false,
                                    error: '無法從回應中找到出勤資料'
                                });
                            }
                            
                        } catch (error) {
                            resolve({
                                success: false,
                                error: '解析回應失敗: ' + error.message
                            });
                        }
                    } else {
                        resolve({
                            success: false,
                            error: `HTTP ${xhr.status}: ${xhr.statusText}`
                        });
                    }
                }
            };
            
            xhr.onerror = function() {
                resolve({
                    success: false,
                    error: '網路錯誤或請求被阻擋'
                });
            };
            
            xhr.ontimeout = function() {
                resolve({
                    success: false,
                    error: '請求超時'
                });
            };
            
            // 設定超時時間
            xhr.timeout = 15000; // 15 秒超時（比較大的資料集）
            
            xhr.send(params.toString());
        });
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // 在頁面載入時檢查是否為目標網站
    if (window.location.hostname === 'cloud.nueip.com') {
        console.log('補卡助手已載入，準備就緒');
    }

})();