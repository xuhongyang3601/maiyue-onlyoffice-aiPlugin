((window, undefined) => {
    // 隐藏浮动窗口
    function hideFloatingWindow() {
        const existingWindow = window.parent.document.getElementById("ai-flota-window");
        if (existingWindow) {
            existingWindow.remove();
        }
    }
    // 文档选中变化
    function handleSelectionChange() {
        hideFloatingWindow();
        window.Asc.plugin.executeMethod("GetSelectedContent", [], (content) => {
            if (content) {
                createFloatingWindow(content);
            } else {
                hideFloatingWindow()
            }
        });
    }
    // 创建浮动窗口
    function createFloatingWindow(content) {
        setTimeout(()=>{
            let rect = window.parent.document.getElementById("id_target_cursor").getBoundingClientRect();
            let btns = [
                {
                    "label": "AI 润色",
                    "aiType": 1
                },
                {
                    "label": "AI 扩写",
                    "aiType": 2
                },
                {
                    "label": "AI 续写",
                    "aiType": 3
                },
                {
                    "label": "AI 缩写",
                    "aiType": 4
                }
            ];
            const buttonContainer = $(`<div id="ai-flota-window" style='position:absolute;top:${rect.top - rect.height}px;left:${rect.right}px;background-color:#fff;padding:10px;border-radius:8px;z-index:1000'></div>`);
            $.each(btns, (index, btn) => {
                const button = $(`<button style="margin: 5px;">${btn.label}</button>`).on("click", () => {
                    window.parent.parent.postMessage({
                        command: 'aiTabs',
                        data: {
                            doc: content,
                            aiType: btn.aiType
                        }
                    }, "*")
                })
                buttonContainer.append(button);
            });
            // 添加到文档中
            window.parent.document.body.appendChild(buttonContainer[0]);
        },50)
    }
    // 处理文档点击事件
    function handleDocumentClick(event) {
        window.Asc.plugin.executeMethod("GetSelectedContent", [], (content) => {
            if (!content) {
                hideFloatingWindow()
            }
        });
    }
    window.Asc.plugin.init = () => {
        window.parent.Common.Gateway.on('internalcommand', (data) => {
            // const { command } = data;
            // switch (command) {
            //     case 'jumpToPositionByIndex':
            //         selectPositionInTheParagraph(data.data);
            //         break;
            //     default:
            //         break;
            // }
        });
        Asc.plugin.attachEvent("onClick", handleDocumentClick);
        handleSelectionChange();
    };
})(window, undefined);