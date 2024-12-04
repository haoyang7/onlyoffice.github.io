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
                var allBookmarksContent = new Map(); // 存储所有书签的内容
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
                                    var bookmarkText = oRange.GetText();
                                    allBookmarksContent.set(bookmarkName, bookmarkText);
                                } else {
                                    allBookmarksContent.set(bookmarkName, "");
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.error(error)
                }
                return allBookmarksContent;
            }, false, true, function (allBookmarksContent) {
                hideLoading();
                console.log('ok', allBookmarksContent)
                var content = "";
                if (allBookmarksContent && allBookmarksContent.size > 0) {
                    var contentArray = [];
                    for (var [key, value] of allBookmarksContent) {
                        contentArray.push("书签名称: " + key + "\n书签内容: " + value + "\n");
                    }
                    content = contentArray.join('');
                } else {
                    content = "文档中没有书签内容";
                }
                $('#bookmarkContent').text(content);
            })
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