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
                                    allBookmarksContent.set(bookmarkName, bookmarkText);
                                } else {
                                    console.log('No range found for bookmark:', bookmarkName);
                                    allBookmarksContent.set(bookmarkName, "");
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
                console.log('callCommand return:', allBookmarksContent);
                return allBookmarksContent;
            }, false, true, function (allBookmarksContent) {
                hideLoading();
                console.log('Callback received. allBookmarksContent:', allBookmarksContent);

                var content = "";
                if (allBookmarksContent && allBookmarksContent.size > 0) {
                    console.log('There are bookmarks. Preparing content...');
                    var contentArray = [];
                    for (var [key, value] of allBookmarksContent) {
                        console.log('Adding bookmark to content: ', key, value);
                        contentArray.push("书签名称: " + key + "\n书签内容: " + value + "\n");
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