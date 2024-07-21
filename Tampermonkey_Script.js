// ==UserScript==
// @name         小红书HTML数据获取器
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  在小红书后台页面右上角生成一个悬浮窗，手动指定总页数，抓取并下载网页HTML代码。
// @author       Skyler1n
// @match        https://creator.xiaohongshu.com/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // 全局变量用于存储所有页面的HTML
    var allHtmlContent = '';
    var currentPage = 1; // 当前页
    var totalPages = 1; // 总页数

    // 更新状态文本的函数
    function updateStatusText(status) {
        statusElement.textContent = status;
    }

    // 用于检测是否有下一页的函数
    function hasNextPage() {
        // 修改这里的选择器以匹配实际的下一页按钮
        var nextPageButtons = document.querySelectorAll('button.dyn.css-1oqsskg.css-19b83d2:not(.disabled)');
        return nextPageButtons.length > 0;
    }

    // 用于模拟点击下一页的函数
    function clickNextPage() {
        var nextPageButtons = document.querySelectorAll('button.dyn.css-1oqsskg.css-19b83d2:not(.disabled)');
        if (nextPageButtons.length > 0) {
            // 默认点击最后一个按钮，通常这是“下一页”按钮
            nextPageButtons[nextPageButtons.length - 1].click();
        }
    }

    // 用于抓取当前页面HTML的函数
    function grabHtmlContent() {
        // 获取当前页面的HTML代码
        var htmlContent = document.documentElement.outerHTML;
        allHtmlContent += htmlContent + '\n\n<!-- Page Separator -->\n\n';
    }

    // 创建“开始抓取”按钮
    var startGrabButton = document.createElement('button');
    startGrabButton.textContent = '开始获取';

    // 创建状态显示元素
    var statusElement = document.createElement('div');
    statusElement.textContent = '准备就绪，等待执行。';

    // 创建输入页数的文本和输入框
    var pageInputText = document.createElement('div');
    pageInputText.textContent = '获取指定页数：';
    var pageInput = document.createElement('input');
    pageInput.type = 'number';
    pageInput.value = '1';
    pageInput.min = '1';


    // 设置点击事件处理函数
    startGrabButton.onclick = function() {
        totalPages = parseInt(pageInput.value);
        updateStatusText('正在抓取第' + currentPage + '页，共' + totalPages + '页');
        // 开始抓取时，先抓取当前页
        grabHtmlContent();

        // 检查是否有下一页，如果有则翻页并抓取
        var interval = setInterval(function() {
            if (hasNextPage() && currentPage < totalPages) {
                clickNextPage();
                currentPage++;
                // 等待页面加载完成
                setTimeout(function() {
                    grabHtmlContent();
                    updateStatusText('正在抓取第' + currentPage + '页，共' + totalPages + '页');
                    // 如果需要，可以在这里添加额外的延迟以确保数据加载完成
                }, 2000); // 调整这个延迟以适应页面加载速度
            } else {
                clearInterval(interval);
                updateStatusText('已完成，请检查下载文件夹！');
                // 所有页面抓取完成后，提供下载链接
                downloadAllHtmlContent();
            }
        }, 3000); // 调整这个间隔以适应页面加载速度
    };

    // 创建悬浮窗的HTML元素
    var floatWindow = document.createElement('div');
    floatWindow.style.cssText = [
        'position: fixed;',
        'top: 10px;',
        'right: 10px;',
        'z-index: 1000;',
        'background-color: white;',
        'padding: 10px;',
        'border: 1px solid black;',
        'border-radius: 5px;'
    ].join(' ');


    // 将文本和输入框添加到悬浮窗中
    floatWindow.appendChild(statusElement);
    floatWindow.appendChild(pageInputText);
    floatWindow.appendChild(pageInput);
    floatWindow.appendChild(startGrabButton);

    // 将悬浮窗添加到页面中
    document.body.appendChild(floatWindow);

    // 用于下载所有HTML内容的函数
    function downloadAllHtmlContent() {
        // 将所有页面的HTML代码转换为Blob对象
        var blob = new Blob([allHtmlContent], { type: 'text/plain' });

        // 创建一个临时的URL链接用于下载
        var url = URL.createObjectURL(blob);
        var downloadLink = document.createElement('a');
        downloadLink.href = url;
        downloadLink.download = 'video_data.txt'; // 文件名
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(url);
    }
})();
