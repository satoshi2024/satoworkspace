                // --- 阶梯缩放 (Step-down Scaling) 核心代码 开始 ---
                const stepDownDrawImage = (targetCtx, img, sx, sy, sw, sh, dx, dy, dw, dh) => {
                    // 如果源尺寸不到目标尺寸的2倍，直接进行最终绘制，防止过度消耗性能
                    if (sw <= dw * 2 || sh <= dh * 2 || dw <= 0 || dh <= 0) {
                        targetCtx.drawImage(img, sx, sy, sw, sh, dx, dy, dw, dh);
                        return;
                    }
                    
                    // 否则，创建一个尺寸减半的临时画布
                    let halfW = Math.floor(sw / 2);
                    let halfH = Math.floor(sh / 2);
                    let tempCanvas = document.createElement('canvas');
                    tempCanvas.width = halfW;
                    tempCanvas.height = halfH;
                    let tempCtx = tempCanvas.getContext('2d');
                    
                    // 临时画布也必须开启高质量平滑
                    tempCtx.imageSmoothingEnabled = true;
                    tempCtx.imageSmoothingQuality = 'high';
                    
                    // 把原图按 50% 缩放画到临时画布上
                    tempCtx.drawImage(img, sx, sy, sw, sh, 0, 0, halfW, halfH);
                    
                    // 递归调用：把刚刚画好的临时画布作为新的原图，继续往下缩
                    stepDownDrawImage(targetCtx, tempCanvas, 0, 0, halfW, halfH, dx, dy, dw, dh);
                };

                // 提取你原来代码第389-390行的坐标和尺寸逻辑，保持完全一致
                let sx = this.m_ImageViewArea.left;
                let sy = this.m_ImageViewArea.top;
                let sw = this.m_ImageViewArea.right - this.m_ImageViewArea.left;
                let sh = this.m_ImageViewArea.bottom - this.m_ImageViewArea.top;
                
                let dx = this.m_DrawArea.left;
                let dy = this.m_DrawArea.top;
                let dw = this.m_DrawArea.left + this.m_DrawArea.right;
                let dh = this.m_DrawArea.top + this.m_DrawArea.bottom;

                // 使用阶梯算法替代原生的 ctx.drawImage
                stepDownDrawImage(ctx, this.imgOrgCanvas, sx, sy, sw, sh, dx, dy, dw, dh);
                // --- 阶梯缩放 (Step-down Scaling) 核心代码 结束 ---
