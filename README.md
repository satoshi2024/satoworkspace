            const stepDownDrawImage = (targetCtx, img, sx, sy, sw, sh, dx, dy, dw, dh) => {
                // 如果源图片的【原始宽高】不到目标尺寸的 2 倍，直接执行最终的裁剪和绘制
                if (img.width <= dw * 2 || img.height <= dh * 2 || dw <= 0 || dh <= 0) {
                    targetCtx.drawImage(img, sx, sy, sw, sh, dx, dy, dw, dh);
                    return;
                }
                
                // 否则，创建一个长宽各减半的临时画布，把整张源图缩小画进去
                let halfW = Math.floor(img.width / 2);
                let halfH = Math.floor(img.height / 2);
                let tempCanvas = document.createElement('canvas');
                tempCanvas.width = halfW;
                tempCanvas.height = halfH;
                let tempCtx = tempCanvas.getContext('2d');
                
                tempCtx.imageSmoothingEnabled = true;
                tempCtx.imageSmoothingQuality = 'high';
                
                // 把【整张】源图缩小一半画到临时画布上
                tempCtx.drawImage(img, 0, 0, img.width, img.height, 0, 0, halfW, halfH);
                
                // 因为图片整体缩小了一半，所以裁剪坐标（sx, sy）和裁剪尺寸（sw, sh）也要同比缩小一半
                let nextSx = sx / 2;
                let nextSy = sy / 2;
                let nextSw = sw / 2;
                let nextSh = sh / 2;
                
                // 递归调用：使用临时画布作为新的源图，并传入缩小后的裁剪参数
                stepDownDrawImage(targetCtx, tempCanvas, nextSx, nextSy, nextSw, nextSh, dx, dy, dw, dh);
            };
