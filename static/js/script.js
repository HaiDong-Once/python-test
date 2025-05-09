document.addEventListener('DOMContentLoaded', function() {
    console.log('脚本加载完成'); // 调试用
    
    // 获取元素
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const uploadForm = document.getElementById('upload-form');
    const uploadLoading = document.getElementById('upload-loading');
    const convertBtns = document.querySelectorAll('.convert-btn');
    const convertLoading = document.getElementById('convert-loading');
    const flashMessages = document.querySelectorAll('.flash-message');
    const deleteButtons = document.querySelectorAll('.confirm-delete');
    
    // 拖放上传功能
    if (uploadArea && fileInput) {
        console.log('初始化文件上传功能'); // 调试用
        
        // 点击上传区域时触发文件选择
        uploadArea.addEventListener('click', function(e) {
            console.log('点击上传区域');
            // 避免重复触发，fileInput已设为覆盖整个区域
            if (e.target === uploadArea) {
                fileInput.click();
            }
        });
        
        // 显示选中的文件名
        fileInput.addEventListener('change', function() {
            console.log('文件选择变化'); // 调试用
            if (this.files.length > 0) {
                const fileName = this.files[0].name;
                const fileNameElem = uploadArea.querySelector('p');
                if (fileNameElem) {
                    fileNameElem.textContent = `已选择: ${fileName}`;
                    console.log(`已选择文件: ${fileName}`); // 调试用
                }
                uploadArea.classList.add('file-selected');
            }
        });
        
        // 拖放文件功能
        uploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', function() {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            if (e.dataTransfer.files.length > 0) {
                fileInput.files = e.dataTransfer.files;
                const fileName = e.dataTransfer.files[0].name;
                const fileNameElem = uploadArea.querySelector('p');
                if (fileNameElem) {
                    fileNameElem.textContent = `已选择: ${fileName}`;
                    console.log(`已选择文件(拖放): ${fileName}`); // 调试用
                }
                uploadArea.classList.add('file-selected');
                
                // 自动触发表单提交
                if (uploadForm) {
                    uploadForm.submit();
                }
            }
        });
    }
    
    // 文件上传进度指示
    if (uploadForm && uploadLoading) {
        uploadForm.addEventListener('submit', function() {
            const fileInput = document.getElementById('file-input');
            if (fileInput && fileInput.files.length > 0) {
                uploadForm.style.display = 'none';
                uploadLoading.style.display = 'block';
                console.log('开始上传文件'); // 调试用
            }
        });
    }
    
    // 标签页转换选项的处理
    if (convertBtns.length > 0 && convertLoading) {
        // 对每个转换按钮添加点击高亮效果
        convertBtns.forEach(function(btn) {
            // 点击转换按钮时
            btn.addEventListener('click', function(e) {
                const filename = this.getAttribute('data-filename');
                if (filename) {
                    convertLoading.style.display = 'block';
                    console.log(`开始转换文件: ${filename}, 使用方法: ${this.getAttribute('data-method')}`);
                    
                    // 高亮当前选中的标签
                    const tabsContainer = this.closest('.tabs-container');
                    if (tabsContainer) {
                        tabsContainer.querySelectorAll('.tab-item').forEach(function(tab) {
                            tab.classList.remove('active');
                        });
                        this.classList.add('active');
                    }
                }
            });
            
            // 悬停效果
            btn.addEventListener('mouseenter', function() {
                this.style.backgroundColor = '#e9ecef';
            });
            
            btn.addEventListener('mouseleave', function() {
                if (!this.classList.contains('active')) {
                    this.style.backgroundColor = '#f8f9fa';
                }
            });
        });
    }
    
    // 删除确认
    if (deleteButtons.length > 0) {
        deleteButtons.forEach(function(btn) {
            btn.addEventListener('click', function(e) {
                e.preventDefault();
                const filename = this.getAttribute('data-filename');
                if (filename && !confirm(`确定要删除文件 ${filename} 吗？`)) {
                    return false;
                }
                console.log(`确认删除文件: ${filename}`); // 调试用
                window.location.href = this.getAttribute('href');
            });
        });
    }
    
    // 闪现消息关闭功能
    if (flashMessages.length > 0) {
        flashMessages.forEach(function(message) {
            const closeBtn = message.querySelector('.close-btn');
            if (closeBtn) {
                closeBtn.addEventListener('click', function() {
                    message.remove();
                });
            }
            
            // 5秒后自动关闭
            setTimeout(function() {
                message.style.opacity = '0';
                message.style.transform = 'translateX(50px)';
                setTimeout(function() {
                    message.remove();
                }, 300);
            }, 5000);
        });
    }
    
    // 添加CSS类以支持动画
    document.head.insertAdjacentHTML('beforeend', `
        <style>
            .dragover {
                border-color: var(--primary-color) !important;
                background-color: #f0f7ff !important;
            }
            
            .file-selected {
                background-color: #f0f7ff;
                border-color: var(--primary-color);
            }
            
            .flash-message {
                transition: opacity 0.3s ease, transform 0.3s ease;
            }
        </style>
    `);
}); 