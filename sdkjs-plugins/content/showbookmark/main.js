(function (window, undefined) {

    // 显示loading
    function showLoading() {
        $('#loading').addClass('show');  // 显示 loading
    }

    // 隐藏loading
    function hideLoading() {
        $('#loading').removeClass('show');  // 隐藏 loading
    }

    window.Asc.plugin.init = function (initData) {
        var me = this
        // 确保在页面加载完成后执行
        $(document).ready(function () {
            showLoading();
            // 官方提供的回调函数，所有操作文档的 API 都可以在这里面使用
            me.callCommand(function () {
                var allBookmarksContent = {}; // 使用普通对象替换 Map
                try {
                    console.log('Calling Api.GetDocument()');
                    var oDocument = Api.GetDocument();
                    if (oDocument) {
                        console.log('Document fetched successfully:', oDocument);
                        var aBookmarks = oDocument.GetAllBookmarksNames();
                        console.log('Fetched bookmarks:', aBookmarks); // 打印获取到的书签名称数组

                        if (aBookmarks && aBookmarks.length > 0) {
                            console.log('Found bookmarks:', aBookmarks.length);
                            for (let i = 0; i < aBookmarks.length; i++) {
                                var bookmarkName = aBookmarks[i];
                                console.log('Processing bookmark:', bookmarkName);
                                var oRange = oDocument.GetBookmarkRange(bookmarkName);
                                console.log('Bookmark range for ' + bookmarkName + ':', oRange);

                                // 如果范围有效，则获取书签内容
                                if (oRange) {
                                    var bookmarkText = oRange.GetText();
                                    console.log('Bookmark text for ' + bookmarkName + ':', bookmarkText);
                                    allBookmarksContent[bookmarkName] = bookmarkText; // 使用对象的方式存储键值对
                                } else {
                                    console.log('No range found for bookmark:', bookmarkName);
                                    allBookmarksContent[bookmarkName] = ""; // 设置为空字符串
                                }
                            }
                        } else {
                            console.log('No bookmarks found.');
                        }
                    } else {
                        console.log('Failed to fetch document.');
                    }
                } catch (error) {
                    console.error('Error in fetching document or processing bookmarks:', error);
                }

                // 转换为 JSON 字符串返回
                var ret = JSON.stringify(allBookmarksContent);
                console.log('callCommand return:', ret);
                return ret;
            }, false, true, function (res) {
                hideLoading();
                console.log('Callback received. res:', res);

                // 解析返回的 JSON 字符串为普通对象
                var allBookmarksContent = JSON.parse(res);
                console.log('Callback allBookmarksContent:', allBookmarksContent);

                var content = "";
                if (allBookmarksContent && Object.keys(allBookmarksContent).length > 0) { // 使用 Object.keys 检查对象是否有内容
                    console.log('There are bookmarks. Preparing content...');
                    var contentArray = [];
                    for (var key in allBookmarksContent) { // 使用 for-in 遍历对象的属性
                        if (allBookmarksContent.hasOwnProperty(key)) {
                            console.log('Adding bookmark to content: ', key, allBookmarksContent[key]);
                            contentArray.push("书签名称: " + key + "\n书签内容: " + allBookmarksContent[key] + "\n");
                        }
                    }
                    content = contentArray.join('');
                } else {
                    console.log('No bookmarks content available.');
                    content = "文档中没有书签内容";
                }

                console.log('Final content:', content);
                $('#bookmarkContent').text(content);
            });
        });

        // 在插件 iframe 之外释放鼠标按钮时调用的函数
        window.Asc.plugin.onExternalMouseUp = function () {
            var event = document.createEvent('MouseEvents')
            event.initMouseEvent('mouseup', true, true, window, 1, 0, 0, 0, 0, false, false, false, false, 0, null)
            document.dispatchEvent(event)
        }

        window.Asc.plugin.button = function (id) {
            // 被中断或关闭窗口
            if (id === -1) {
                this.executeCommand('close', '')
            }
        }
    }
})(window, undefined)