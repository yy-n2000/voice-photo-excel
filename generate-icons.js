// 运行: node generate-icons.js
const { createCanvas } = require('canvas');
const fs = require('fs');

function makeIcon(size) {
  const canvas = createCanvas(size, size);
  const ctx    = canvas.getContext('2d');

  // 背景
  const grad = ctx.createLinearGradient(0, 0, size, size);
  grad.addColorStop(0, '#4472C4');
  grad.addColorStop(1, '#2d5aa0');
  ctx.fillStyle = grad;
  ctx.roundRect(0, 0, size, size, size * 0.2);
  ctx.fill();

  // 麦克风图标文字
  ctx.fillStyle = 'white';
  ctx.font      = `bold ${size * 0.45}px Arial`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('🎤', size / 2, size / 2);

  return canvas.toBuffer('image/png');
}

fs.writeFileSync('public/icon-192.png', makeIcon(192));
fs.writeFileSync('public/icon-512.png', makeIcon(512));
console.log('图标生成完成！');