(function(window, undefined) {
    window.Asc.plugin.init = function(initData) {
      var me = this
      $('#showBookmark').click(function() {
        // 官方提供的回调函数，所有操作文档的 API 都可以在这里面使用
        me.callCommand(function() {
          try {
            var oDocument = Api.GetDocument();
            if (oDocument) {
              var aBookmarks = oDocument.GetAllBookmarksNames();
              if (aBookmarks) {
                for (let i = 0; i < aBookmarks.length; i++) {
                  var oRange = oDocument.GetBookmarkRange(aBookmarks[i]);
                  console.log("bookmark: " + i + ", range:", oRange);
                }
              }
            }
          } catch (error) {
            console.error(error)
          }
        }, false, true, function () {
          console.log('ok')
        })
      })
  
      // 在插件 iframe 之外释放鼠标按钮时调用的函数
      window.Asc.plugin.onExternalMouseUp = function() {
        var event = document.createEvent('MouseEvents')
        event.initMouseEvent('mouseup', true, true, window, 1, 0, 0, 0, 0, false, false, false, false, 0, null)
        document.dispatchEvent(event)
      }
  
      window.Asc.plugin.button = function(id) {
        // 被中断或关闭窗口
        if (id === -1) {
          this.executeCommand('close', '')
        }
        }
    }
  })(window, undefined)