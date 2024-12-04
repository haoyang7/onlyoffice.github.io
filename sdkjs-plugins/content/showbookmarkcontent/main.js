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
                var allBookmarksContent = {};
                try {
                    var oDocument = Api.GetDocument();
                    if (oDocument) {
                        var aBookmarks = oDocument.GetAllBookmarksNames();

                        if (aBookmarks && aBookmarks.length > 0) {
                            for (let i = 0; i < aBookmarks.length; i++) {
                                var bookmarkName = aBookmarks[i];
                                var oRange = oDocument.GetBookmarkRange(bookmarkName);

                                // 如果范围有效，则获取书签内容
                                if (oRange) {
                                    allBookmarksContent[bookmarkName] = oRange.GetText();
                                } else {
                                    allBookmarksContent[bookmarkName] = ""; // 设置为空字符串
                                }
                            }
                        } else {
                            console.error('No bookmarks found.');
                        }
                    } else {
                        console.error('Failed to fetch document.');
                    }
                } catch (error) {
                    console.error('Error in fetching document or processing bookmarks:', error);
                }

                // 转换为 JSON 字符串返回
                return JSON.stringify(allBookmarksContent);
            }, false, true, function (res) {
                hideLoading();

                // 解析返回的 JSON 字符串
                var allBookmarksContent = JSON.parse(res);

                var content = "";
                if (allBookmarksContent && Object.keys(allBookmarksContent).length > 0) { // 使用 Object.keys 检查对象是否有内容
                    var contentArray = [];
                    for (var key in allBookmarksContent) { // 使用 for-in 遍历对象的属性
                        if (allBookmarksContent.hasOwnProperty(key)) {
                            contentArray.push("书签名称: " + key + "\n书签内容: " + allBookmarksContent[key] + "\n");
                        }
                    }
                    content = contentArray.join('');
                } else {
                    content = "文档中没有书签内容";
                }

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