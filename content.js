// 将图片转换为base64格式，保持网页显示大小
async function convertImagesToBase64(container) {
  const images = container.getElementsByTagName('img');
  
  for (let img of images) {
    if (img.src.startsWith('data:')) continue;
    
    try {
      const response = await fetch(img.src);
      const blob = await response.blob();
      const reader = new FileReader();
      
      await new Promise((resolve, reject) => {
        reader.onload = () => {
          // 获取图片在网页中实际显示的大小
          const computedStyle = window.getComputedStyle(img);
          const displayWidth = parseFloat(computedStyle.width);
          const displayHeight = parseFloat(computedStyle.height);
          
          // 设置图片属性，使用网页中显示的大小
          img.src = reader.result;
          img.width = displayWidth;
          img.height = displayHeight;
          img.style.width = displayWidth + 'px';
          img.style.height = displayHeight + 'px';
          
          // 移除可能影响大小的样式
          img.style.maxWidth = 'none';
          img.style.maxHeight = 'none';
          resolve();
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    } catch (error) {
      console.error('图片转换失败:', error);
      img.remove();
    }
  }
}

// 添加样式
function addStyles(container) {
  // 添加基本样式
  const style = document.createElement('style');
  style.textContent = `
    body { font-family: Arial, sans-serif; }
    p { margin: 10px 0; }
    img { 
      /* 保持图片大小，不进行缩放 */
      max-width: none !important;
      max-height: none !important;
      width: auto !important;
      height: auto !important;
    }
    table { border-collapse: collapse; width: 100%; }
    td, th { border: 1px solid black; padding: 8px; }
    h1 { font-size: 24px; }
    h2 { font-size: 20px; }
    h3 { font-size: 16px; }
  `;
  container.appendChild(style);
}

// 主要保存函数
async function saveAsWord() {
  // 获取选中内容
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    showNotification('请先选择要保存的内容！', 'error');
    return;
  }

  // 显示加载提示
  showNotification('正在处理内容...', 'info');

  try {
    // 创建临时容器
    const tempContainer = document.createElement('div');
    const range = selection.getRangeAt(0);
    tempContainer.appendChild(range.cloneContents());

    // 处理图片
    await convertImagesToBase64(tempContainer);

    // 添加样式
    addStyles(tempContainer);

    // 转换为Word文档
    const content = tempContainer.innerHTML;
    const converted = htmlDocx.asBlob(content, {
      orientation: 'portrait',
      margins: { top: 720, right: 720, bottom: 720, left: 720 },
      // 指定生成docx格式
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    // 设置文件名（改为.docx扩展名）
    const timestamp = new Date().toISOString().slice(0,19).replace(/[:]/g, '-');
    const pageTitle = document.title.replace(/[<>:"/\\|?*]/g, '-') || 'webpage';
    const fileName = `${pageTitle}_${timestamp}.docx`;

    // 下载文件
    const url = URL.createObjectURL(converted);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    showNotification('文档保存成功！', 'success');
  } catch (error) {
    console.error('保存失败:', error);
    showNotification('保存失败，请重试', 'error');
  }
}

// 显示通知
function showNotification(message, type) {
  const colors = {
    success: '#4CAF50',
    error: '#F44336',
    info: '#2196F3'
  };
  
  const notification = document.createElement('div');
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: ${colors[type]};
    color: white;
    padding: 15px 25px;
    border-radius: 5px;
    z-index: 9999;
    box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    font-family: Arial, sans-serif;
    font-size: 14px;
    transition: opacity 0.3s ease-in-out;
  `;
  notification.textContent = message;
  document.body.appendChild(notification);
  
  setTimeout(() => {
    notification.style.opacity = '0';
    setTimeout(() => {
      document.body.removeChild(notification);
    }, 300);
  }, 2700);
}

// 执行保存功能
saveAsWord();