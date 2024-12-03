(function (window, undefined) {
    window.Asc.plugin.init = function () {

    }

    // 确保在页面加载完成后执行
    $(document).ready(function () {
        $('#showBookmark').click(function () {
            var allBookmarksContent = ""; // 存储所有书签的内容
            var aBookmarks;
            window.Asc.plugin.executeMethod('GetAllBookmarksNames', [], function (sOutput) {
                aBookmarks = sOutput;
            });
            if (aBookmarks && aBookmarks.length > 0) {
                for (let i = 0; i < aBookmarks.length; i++) {
                    var bookmarkName = aBookmarks[i];
                    var oRange;
                    window.Asc.plugin.executeMethod('GetBookmarkRange', [bookmarkName], function (sOutput) {
                        oRange = sOutput;
                    });
                    // 如果范围有效，则获取书签内容
                    if (oRange) {
                        var bookmarkText = oRange.GetText();
                        allBookmarksContent += "书签名称: " + bookmarkName + "<br>书签内容: " + bookmarkText + "<br><br>"; // 将每个书签的内容加入到字符串中
                    } else {
                        allBookmarksContent += "书签名称: " + bookmarkName + " 的范围未找到<br><br>";
                    }
                }
            }
            console.log("书签内容：", allBookmarksContent);
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
})(window, undefined)