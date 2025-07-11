((window, undefined) => {
  function markdownToHtml(mdString) {
    // 配置选项（按需启用）
    marked.setOptions({
      gfm: true,
      breaks: true, // 转换换行符
      sanitize: false,
    });

    return marked.parse(mdString);
  }
  // 隐藏浮动窗口
  function hideFloatingWindow() {
    const existingWindow =
      window.parent.document.getElementById("ai-flota-window");
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
        hideFloatingWindow();
      }
    });
  }
  // 更新浮动窗口位置
  function updateFloatingWindowPosition() {
    const floatingWindow = window.parent.document.getElementById("ai-flota-window");
    if (floatingWindow) {
      let rect = window.parent.document
        .getElementById("id_target_cursor")
        .getBoundingClientRect();
      floatingWindow.style.top = `${rect.top - rect.height - 30}px`;
      floatingWindow.style.left = `${rect.right}px`;
    }
  }
  // 处理目标位置变化事件
  function handleTargetPositionChanged() {
    window.Asc.plugin.executeMethod("GetSelectedContent", [], (content) => {
      if (content) {
        // 更新浮动窗口位置
        if (window.parent.document
          .getElementById("ai-flota-window")) {
          updateFloatingWindowPosition();
        }
      } else {
        hideFloatingWindow();
      }
    });
  }
  // 创建浮动窗口
  function createFloatingWindow(content) {
    setTimeout(() => {
      let rect = window.parent.document
        .getElementById("id_target_cursor")
        .getBoundingClientRect();
      let btns = [
        {
          label: "AI 润色",
          aiType: 1,
        },
        {
          label: "AI 扩写",
          aiType: 2,
        },
        {
          label: "AI 续写",
          aiType: 3,
        },
        {
          label: "AI 缩写",
          aiType: 4,
        },
      ];
      const buttonContainer = $(
        `<div id="ai-flota-window" style='position:absolute;top:${rect.top - rect.height
        }px;left:${rect.right
        }px;background-color:#fff;padding:10px;border-radius:8px;z-index:1000'></div>`
      );
      $.each(btns, (index, btn) => {
        const button = $(
          `<button style="margin: 5px;">${btn.label}</button>`
        ).on("click", () => {
          window.parent.parent.postMessage(
            {
              command: "aiTabs",
              frameEditorId: window.parent.frameEditorId,
              data: {
                doc: content,
                aiType: btn.aiType,
              },
            },
            "*"
          );
        });
        buttonContainer.append(button);
      });
      // 添加到文档中
      window.parent.document.body.appendChild(buttonContainer[0]);
    }, 50);
  }
  // 插入 AI 内容
  function insertAiContent(content) {
    clearSelectionOverlay();
    // window.Asc.scope.insertAiContentData = markdownToHtml(content);
    window.Asc.plugin.callCommand(
      () => {
        let doc = Api.GetDocument();
        // console.log(Asc)
        // let paragraph1 = Api.CreateParagraph();
        let paragraph2 = Api.CreateParagraph();
        const defaultStyles = {
          bold: false,
          italic: false,
          underline: false,
          fontSize: 16, // 默认字号
          fontFamily: "Calibri", // 默认字体
        };

        // 应用默认样式
        paragraph2.SetBold(defaultStyles.bold);
        paragraph2.SetItalic(defaultStyles.italic);
        paragraph2.SetUnderline(defaultStyles.underline);
        paragraph2.SetFontSize(defaultStyles.fontSize);
        paragraph2.SetFontFamily(defaultStyles.fontFamily);
        paragraph2.SetColor(255, 255, 255);
        paragraph2.AddText(".");
        // doc.AddElement(0, paragraph2);
        doc.Push(paragraph2);
        let range = Api.CreateRange(doc, 0);
        range.Select();
        range.SetHighlight("white");
        // doc.RemoveAllElements();
      },
      false,
      false,
      (endPos) => {
        window.Asc.plugin.executeMethod(
          "PasteHtml",
          [markdownToHtml(content)],
          (x) => {
            window.parent.parent.postMessage(
              {
                command: "insertAiContent",
                frameEditorId: window.parent.frameEditorId,
              },
              "*"
            );
          }
        );
      }
    );
  }
  //   显示selection蒙层
  function showSelectionOverlay() {
    if (window.parent.document.getElementById("onlyoffice-selection-hider")) {
      window.parent.document
        .getElementById("onlyoffice-selection-hider")
        .remove();
      window.parent.parent.postMessage(
        {
          command: "showSelectionOverlay",
          frameEditorId: window.parent.frameEditorId,
        },
        "*"
      );
    }
  }
  //  隐藏selection蒙层
  function clearSelectionOverlay() {
    if (!window.parent.document.getElementById("onlyoffice-selection-hider")) {
      const css = `
          #id_viewer_overlay{
              display:none !important;
          }
      `;
      const style = document.createElement("style");
      style.id = "onlyoffice-selection-hider";
      style.textContent = css;
      window.parent.document.head.appendChild(style);
    }
  }
  window.Asc.plugin.init = () => {
    window.parent.Common.Gateway.on("internalcommand", (data) => {
      const { command } = data;
      switch (command) {
        case "insertAiContent":
          insertAiContent(data.data);
          break;
        // 显示selection蒙层
        case "showSelectionOverlay":
          showSelectionOverlay();
          break;
        default:
          break;
      }
    });
    Asc.plugin.attachEvent("onTargetPositionChanged", handleTargetPositionChanged);
    handleSelectionChange();
  };
})(window, undefined);
