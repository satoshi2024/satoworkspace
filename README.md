    DrawImage = () => {
        let ctx = this.uiLayer.getContext('2d');
        ctx.clearRect(0, 0, this.uiLayer.width, this.uiLayer.height);
        this.DrawBgColor();

        // ========================== 【新增开始】========================== 
        // 阶梯缩放辅助函数：专门解决 Canvas 剧烈缩小导致像素断裂变糊的问题
        const stepDownDrawImage = (targetCtx, img, sx, sy, sw, sh, dx, dy, dw, dh) => {
            // 如果缩放比例在 2 倍以内，或者尺寸异常，直接交由原生处理
            if (sw <= dw * 2 || sh <= dh * 2 || dw <= 0 || dh <= 0) {
                targetCtx.imageSmoothingEnabled = true;
                targetCtx.imageSmoothingQuality = 'high';
                targetCtx.drawImage(img, sx, sy, sw, sh, dx, dy, dw, dh);
                return;
            }
            
            // 否则，按 50% 比例创建临时离屏画布进行逐步缩小
            let halfW = Math.max(Math.floor(sw / 2), dw);
            let halfH = Math.max(Math.floor(sh / 2), dh);
            let tempCanvas = document.createElement('canvas');
            tempCanvas.width = halfW;
            tempCanvas.height = halfH;
            let tempCtx = tempCanvas.getContext('2d');
            tempCtx.imageSmoothingEnabled = true;
            tempCtx.imageSmoothingQuality = 'high';
            
            // 将原图画到一半大小的临时画布上
            tempCtx.drawImage(img, sx, sy, sw, sh, 0, 0, halfW, halfH);
            
            // 递归调用，直到尺寸足够小
            stepDownDrawImage(targetCtx, tempCanvas, 0, 0, halfW, halfH, dx, dy, dw, dh);
        };
        // ========================== 【新增结束】========================== 

        if (this.m_iStatus == CommonDefine.SHOW_IMAGE) {
            if (this.imgOrgData != null && this.imgOrgCanvas != null) {
                
                // 将原来粗暴的 ctx.drawImage 替换为阶梯缩放调用
                // (变量值完全保持你原代码 386-388 行的参数逻辑)
                let sx = this.m_ImageViewArea.left;
                let sy = this.m_ImageViewArea.top;
                let sw = this.m_ImageViewArea.right - this.m_ImageViewArea.left;
                let sh = this.m_ImageViewArea.bottom - this.m_ImageViewArea.top;
                let dx = this.m_DrawArea.left;
                let dy = this.m_DrawArea.top;
                let dw = this.m_DrawArea.left + this.m_DrawArea.right; 
                let dh = this.m_DrawArea.top + this.m_DrawArea.bottom;

                // 执行平滑渲染
                stepDownDrawImage(ctx, this.imgOrgCanvas, sx, sy, sw, sh, dx, dy, dw, dh);

                // --- 原代码 389 行及以后的虚线/框线绘制逻辑保留不变 ---
                if (this.m_bDrawedRect && this.m_bDrawedRectPos) {
