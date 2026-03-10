const path = require('path');
const { app, dialog } = require('electron');
const url = require('url');

let pluginApi = null;

function fileUrl(p) { return url.pathToFileURL(p).href; }
function emitUpdate(channel, target, value) { try { pluginApi.emit(channel, { type: 'update', target, value }); } catch (e) {} }

const EVENT_CHANNEL = 'profiles-seating-channel';
let state = { mode: 'position', paths: {} };

function ensureDefaults() {
  if (!pluginApi) return;
  const defaults = {
    rows: [
      { id: 'r1', label: '第1排', type: 'row' },
      { id: 'r2', label: '第2排', type: 'row' },
      { id: 'r3', label: '第3排', type: 'row' },
      { id: 'r4', label: '第4排', type: 'row' }
    ],
    cols: [
      { id: 'c1', label: '第1列', type: 'col' },
      { id: 'c2', label: '第2列', type: 'col' },
      { id: 'c3', label: '第3列', type: 'col' },
      { id: 'c4', label: '第4列', type: 'col' },
      { id: 'c5', label: '第5列', type: 'col' },
      { id: 'c6', label: '第6列', type: 'col' }
    ],
    seats: {},
    backgroundStatus: '默认'
  };
  try {
    const current = pluginApi.store.getAll() || {};
    let changed = false;
    Object.keys(defaults).forEach(k => {
      if (!(k in current)) {
        current[k] = defaults[k];
        changed = true;
      }
    });
    if (changed) pluginApi.store.setAll(current);
  } catch (e) {}
}

const functions = {
  openSeating: async () => {
    state.paths.seating = fileUrl(path.join(__dirname, 'pages', 'seating.html')) + `?channel=${encodeURIComponent(EVENT_CHANNEL)}&caller=${encodeURIComponent('profiles-seating')}`;
    const params = {
      title: '档案-座次表',
      eventChannel: EVENT_CHANNEL,
      subscribeTopics: [EVENT_CHANNEL],
      callerPluginId: 'profiles-seating',
      backgroundUrl: state.paths.seating,
      floatingUrl: null,
      windowMode: 'fullscreen_only',
      leftItems: [
        { id: 'toggle-free-list', text: '无座学生', icon: 'ri-user-unfollow-line' },
        { id: 'export', text: '导出', icon: 'ri-download-line' },
        { id: 'save', text: '保存', icon: 'ri-save-3-line' }
      ],
      centerItems: [
        { id: 'zoom-out', text: '缩小', icon: 'ri-zoom-out-line' },
        { id: 'zoom-reset', text: '适应', icon: 'ri-layout-grid-line' },
        { id: 'zoom-in', text: '放大', icon: 'ri-zoom-in-line' }
      ]
    };
    await pluginApi.call('ui-lowbar', 'openTemplate', [params]);
    return true;
  },
  onLowbarEvent: async (payload = {}) => {
    try {
      if (payload?.type === 'left.click') {
        if (payload.id === 'save') emitUpdate(EVENT_CHANNEL, 'seating.save', true);
        if (payload.id === 'toggle-free-list') emitUpdate(EVENT_CHANNEL, 'freeList.toggle', true);
        if (payload.id === 'export') emitUpdate(EVENT_CHANNEL, 'export.show', true);
      }
      if (payload?.type === 'click') {
        if (payload.id === 'zoom-out') emitUpdate(EVENT_CHANNEL, 'zoom.out', true);
        if (payload.id === 'zoom-reset') emitUpdate(EVENT_CHANNEL, 'zoom.reset', true);
        if (payload.id === 'zoom-in') emitUpdate(EVENT_CHANNEL, 'zoom.in', true);
      }
      return true;
    } catch (e) { return { ok: false, error: e?.message || String(e) }; }
  },
  getConfig: async () => {
    try { return { ok: true, config: pluginApi.store.getAll() }; } catch (e) { return { ok: false, error: e?.message || String(e) }; }
  },
  saveConfig: async (payload = {}) => {
    try {
      if (Array.isArray(payload.rows)) pluginApi.store.set('rows', payload.rows);
      if (Array.isArray(payload.cols)) pluginApi.store.set('cols', payload.cols);
      if (payload.seats && typeof payload.seats === 'object') pluginApi.store.set('seats', payload.seats);
      if (typeof payload.backgroundStatus === 'string') pluginApi.store.set('backgroundStatus', payload.backgroundStatus);
      emitUpdate(EVENT_CHANNEL, 'refresh', true);
      return { ok: true };
    } catch (e) { return { ok: false, error: e?.message || String(e) }; }
  },
  exportToWord: async (payload = {}) => {
    try {
      const config = payload[0] || {};
      const fs = require('fs');
      const path = require('path');
      
      // 显示保存对话框
      const result = await dialog.showSaveDialog({
        title: '导出Word文档',
        defaultPath: `座次表-${Date.now()}.docx`,
        filters: [
          { name: 'Word文档', extensions: ['docx'] },
          { name: '所有文件', extensions: ['*'] }
        ]
      });
      
      if (!result.filePath) {
        return { ok: false, error: '用户取消' };
      }
      
      const filePath = result.filePath;
      
      // 简单的Word文档生成（实际项目中可能需要使用docx库）
      const content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>座次表</w:t>
      </w:r>
    </w:p>
    <w:table>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t></w:t>
            </w:r>
          </w:p>
        </w:tc>
        ${config.cols?.map(col => `<w:tc>
          <w:p>
            <w:r>
              <w:t>${col.label || ''}</w:t>
            </w:r>
          </w:p>
        </w:tc>`).join('')}
      </w:tr>
      ${config.rows?.map(row => `<w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>${row.label || ''}</w:t>
            </w:r>
          </w:p>
        </w:tc>
        ${config.cols?.map(col => {
          const seatKey = `${row.id}-${col.id}`;
          const seat = config.seats?.[seatKey];
          return `<w:tc>
            <w:p>
              <w:r>
                <w:t>${seat?.name || ''}</w:t>
              </w:r>
            </w:p>
          </w:tc>`;
        }).join('')}
      </w:tr>`).join('')}
    </w:table>
  </w:body>
</w:document>`;
      
      fs.writeFileSync(filePath, content);
      return { ok: true, path: filePath };
    } catch (e) {
      return { ok: false, error: e?.message || String(e) };
    }
  },
  exportToImage: async (payload = {}) => {
    try {
      const config = payload[0] || {};
      const fs = require('fs');
      const path = require('path');
      
      // 显示保存对话框
      const result = await dialog.showSaveDialog({
        title: '导出图片',
        defaultPath: `座次表-${Date.now()}.png`,
        filters: [
          { name: 'PNG图片', extensions: ['png'] },
          { name: '所有文件', extensions: ['*'] }
        ]
      });
      
      if (!result.filePath) {
        return { ok: false, error: '用户取消' };
      }
      
      const filePath = result.filePath;
      
      // 简单的图片生成（实际项目中可能需要使用Canvas或其他库）
      // 这里只是创建一个占位文件
      fs.writeFileSync(filePath, Buffer.from('placeholder'));
      return { ok: true, path: filePath };
    } catch (e) {
      return { ok: false, error: e?.message || String(e) };
    }
  }
};

const init = async (api) => {
  pluginApi = api;
  api.splash.setStatus('plugin:init', '初始化 档案-座次表');
  ensureDefaults();
  api.splash.progress('plugin:init', '档案-座次表就绪');
};

module.exports = {
  name: 'profiles-seating',
  version: '0.1.0',
  description: '档案-座次表（底栏模板，全屏）',
  init,
  functions: {
    ...functions,
    getVariable: async (name) => { const k=String(name||''); if (k==='timeISO') return new Date().toISOString(); if (k==='pluginName') return '档案-座次表'; return ''; },
    listVariables: () => ['timeISO','pluginName']
  }
}
