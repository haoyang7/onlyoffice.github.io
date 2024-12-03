(function (window, undefined) {
    window.Asc.plugin.init = function (initData) {
        var me = this
        $(document).ready(function () {
            $('#showBookmark').click(function () {
                // 官方提供的回调函数，所有操作文档的 API 都可以在这里面使用
                me.callCommand(function () {
                    try {
                        var oDocument = Api.GetDocument();
                        var allBookmarksContent = ""; // 存储所有书签的内容
                        if (oDocument) {
                            var aBookmarks = oDocument.GetAllBookmarksNames();
                            if (aBookmarks && aBookmarks.length > 0) {
                                for (let i = 0; i < aBookmarks.length; i++) {
                                    var bookmarkName = aBookmarks[i];
                                    var oRange = oDocument.GetBookmarkRange(bookmarkName);
                                    // 如果范围有效，则获取书签内容
                                    if (oRange) {
                                        var bookmarkText = oRange.GetText();
                                        allBookmarksContent += "书签名称: " + bookmarkName + "<br>书签内容: " + bookmarkText + "<br><br>"; // 将每个书签的内容加入到字符串中
                                    } else {
                                        allBookmarksContent += "书签名称: " + bookmarkName + " 的范围未找到<br><br>";
                                    }
                                }

                            }
                        }
                    } catch (error) {
                        console.error(error)
                    }
                }, false, true, function () {
                    console.log('ok')
                })
                // 通过 iframe 的 id 或其他方式访问 iframe
                var iframe = document.getElementById('iframe_asc.{11700c35-1fdb-4e37-9edb-b31637139601}');  // 获取 iframe
                var iframeDocument = iframe.contentWindow.document;  // 获取 iframe 的 document
                // 在 iframe 内部查找元素
                var bookmarkContent = $(iframeDocument).find('#bookmarkContent');
                console.log(bookmarkContent.length);  // 应该输出 1，表示元素存在
                if (allBookmarksContent) {
                    console.log("书签内容：", allBookmarksContent); // 一次性弹出所有书签内容
                    $("#bookmarkContent", iframeDocument).html(allBookmarksContent);
                } else {
                    console.log("文档中没有书签内容");
                    $("#bookmarkContent", iframeDocument).html("文档中没有书签内容");
                }
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