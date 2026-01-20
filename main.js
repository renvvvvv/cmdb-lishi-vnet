import { bitable } from '@lark-base-open/js-sdk';
import { DataAPI, DataTransformer, API_PRESETS, AGGREGATION_FUNCTIONS } from './api.js';

// 全局状态
let state = {
  sourceTable: null,  // 数据源表格（读取测点ID）- 默认当前表格
  targetTable: null,  // 目标表格（写入结果）
  selection: null,
  isSyncing: false,
  isPaused: false,    // 是否暂停
  api: null,
  syncProgress: {     // 同步进度
    currentIndex: 0,  // 当前处理到第几条记录
    recordIdList: [], // 所有记录ID列表
    total: 0,         // 总记录数
    successCount: 0,  // 成功数量
    failCount: 0,     // 失败数量
    isWriting: false, // 当前是否正在写入数据
    currentPointId: null  // 当前正在处理的测点ID
  },
  config: {
    sourceTableUrl: '',     // 数据源表格URL（留空=当前表格）
    sourceTableId: '',      // 数据源表格ID
    targetTableUrl: '',     // 目标表格URL（必填）
    targetTableId: '',      // 目标表格ID
    tableType: '',
    appId: 'cli_a8495c3aba10d00b',  // 默认App ID
    appSecret: '2jyb2v28yXHPeKW2IuIqPlwyigTP66uq',  // 默认App Secret
    apiUrl: 'http://localhost:3001',
    authHeader: 'Basic dGVjaG5pcXVlX2NlbnRlcjoyMVZpYW5ldEBWbmV0LmNvbQ==',
    queryColumn: '测点',
    startTime: '',
    endTime: '',
    interval: '3600',
    aggregationFunction: '',
    dataFormat: 'latest'
  }
};

// 初始化
async function init() {
  try {
    // 获取当前表格（默认作为数据源表格）
    const currentTable = await bitable.base.getActiveTable();
    state.selection = await bitable.base.getSelection();
    
    // 默认使用当前表格作为数据源
    state.sourceTable = currentTable;
    
    // 检测表格类型
    await detectTableType();
    
    // 加载配置
    loadConfig();
    
    // 如果配置了数据源表格URL，尝试获取数据源表格
    if (state.config.sourceTableUrl && state.config.sourceTableId) {
      try {
        log('尝试连接数据源表格...', 'info');
        state.sourceTable = await getTableByUrl(state.config.sourceTableUrl, state.config.sourceTableId);
        const sourceName = await state.sourceTable.getName();
        log(`✓ 已连接到数据源表格: ${sourceName}`, 'success');
      } catch (error) {
        log(`无法连接数据源表格: ${error.message}，将使用当前表格`, 'warning');
        state.sourceTable = currentTable;
      }
    }
    
    // 如果配置了目标表格URL，尝试获取目标表格
    if (state.config.targetTableUrl && state.config.targetTableId) {
      try {
        log('尝试连接目标表格...', 'info');
        state.targetTable = await getTableByUrl(state.config.targetTableUrl, state.config.targetTableId);
        const targetName = await state.targetTable.getName();
        log(`✓ 已连接到目标表格: ${targetName}`, 'success');
      } catch (error) {
        log(`无法连接目标表格: ${error.message}，将使用当前表格`, 'warning');
        state.targetTable = currentTable;
      }
    } else {
      // 没有配置目标表格，使用当前表格
      state.targetTable = currentTable;
      log('目标表格未配置，将使用当前表格', 'info');
    }
    
    // 初始化 API
    state.api = new DataAPI({
      baseUrl: state.config.apiUrl,
      authHeader: state.config.authHeader
    });
    
    // 绑定事件
    bindEvents();
    
    // 加载保存的进度
    const hasProgress = loadProgress();
    if (hasProgress) {
      document.getElementById('resumeSync').disabled = false;
      updateProgress(
        state.syncProgress.currentIndex,
        state.syncProgress.total,
        state.syncProgress.successCount,
        state.syncProgress.failCount
      );
    }
    
    log('历史数据查询插件初始化成功', 'success');
    log(`API 地址: ${state.config.apiUrl}`, 'info');
    log(`使用${state.config.apiUrl.includes('localhost') ? '代理服务器' : '直连'}模式`, 'info');
    
    // 设置默认时间（最近1小时）
    setDefaultTimeRange();
  } catch (error) {
    log(`初始化失败: ${error.message}`, 'error');
  }
}

// 解析表格 URL
function parseTableUrl(url) {
  if (!url) return null;
  
  try {
    // Base 表格格式: https://xxx.feishu.cn/base/xxxxx?table=tblxxxx
    // Wiki 表格格式: https://xxx.feishu.cn/wiki/xxxxx
    
    const urlObj = new URL(url);
    const pathname = urlObj.pathname;
    
    let type = '';
    let tableId = '';
    let baseId = '';
    
    if (pathname.includes('/base/')) {
      type = 'base';
      const matches = pathname.match(/\/base\/([^/?]+)/);
      if (matches) {
        baseId = matches[1];
      }
      // 从 URL 参数获取 table ID
      tableId = urlObj.searchParams.get('table') || '';
    } else if (pathname.includes('/wiki/')) {
      type = 'wiki';
      const matches = pathname.match(/\/wiki\/([^/?]+)/);
      if (matches) {
        tableId = matches[1];
      }
    }
    
    return {
      type,
      tableId,
      baseId,
      isValid: !!(type && (tableId || baseId))
    };
  } catch (error) {
    return null;
  }
}

// 根据URL获取表格对象
async function getTableByUrl(url, tableId) {
  try {
    const parsed = parseTableUrl(url);
    
    if (!parsed || !parsed.isValid) {
      throw new Error('无效的表格URL');
    }
    
    // 获取所有表格
    const tableMetaList = await bitable.base.getTableMetaList();
    console.log('所有表格列表:', tableMetaList);
    
    // 查找匹配的表格
    let targetTableMeta = null;
    
    // 优先使用 tableId 匹配
    if (tableId) {
      targetTableMeta = tableMetaList.find(t => t.id === tableId);
    }
    
    // 如果没找到且有 baseId，尝试匹配 baseId
    if (!targetTableMeta && parsed.baseId) {
      // 注意：baseId 通常是整个多维表格的ID，不是单个表格的ID
      // 这种情况下我们可能需要获取第一个表格
      targetTableMeta = tableMetaList[0];
    }
    
    if (!targetTableMeta) {
      throw new Error(`未找到表格 ID: ${tableId}`);
    }
    
    console.log('找到目标表格:', targetTableMeta);
    
    // 获取表格对象
    const table = await bitable.base.getTable(targetTableMeta.id);
    return table;
    
  } catch (error) {
    console.error('获取表格失败:', error);
    throw error;
  }
}

// 检测表格类型（base 或 wiki）
async function detectTableType() {
  try {
    const baseInfo = await bitable.base.getBaseInfo();
    const type = baseInfo.type || 'base';
    
    state.config.tableType = type;
    
    // 注意：不再设置界面元素，因为现在使用数据源和目标表格分离
    log(`当前Base类型: ${type === 'base' ? 'Base 多维表格' : 'Wiki 多维表格'}`, 'info');
  } catch (error) {
    log('无法检测表格类型', 'warning');
  }
}

// 设置默认时间范围（最近1小时）
function setDefaultTimeRange() {
  const now = new Date();
  const oneHourAgo = new Date(now.getTime() - 3600000);
  
  document.getElementById('startTime').value = formatDateTimeLocal(oneHourAgo);
  document.getElementById('endTime').value = formatDateTimeLocal(now);
}

// 格式化日期时间为 datetime-local 格式
function formatDateTimeLocal(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}`;
}

// 加载配置
function loadConfig() {
  const saved = localStorage.getItem('syncPluginConfig');
  if (saved) {
    const config = JSON.parse(saved);
    
    // 强制检查并修复 API 地址
    if (config.apiUrl && config.apiUrl.includes('digitaltwin.meta42.indc.vnet.com')) {
      console.warn('检测到旧的 API 地址，自动修复为代理服务器地址');
      config.apiUrl = 'http://localhost:3001';
      localStorage.setItem('syncPluginConfig', JSON.stringify(config));
      log('已自动修复 API 地址为代理服务器', 'info');
    }
    
    state.config = { ...state.config, ...config };
  }
  
  // 设置默认值到界面
  document.getElementById('sourceTableUrl').value = state.config.sourceTableUrl || '';
  document.getElementById('targetTableUrl').value = state.config.targetTableUrl || '';
  document.getElementById('appId').value = state.config.appId || 'cli_a8495c3aba10d00b';
  document.getElementById('appSecret').value = state.config.appSecret || '2jyb2v28yXHPeKW2IuIqPlwyigTP66uq';
  document.getElementById('apiUrl').value = state.config.apiUrl || 'http://localhost:3001';
  document.getElementById('authHeader').value = state.config.authHeader || 'Basic dGVjaG5pcXVlX2NlbnRlcjoyMVZpYW5ldEBWbmV0LmNvbQ==';
  document.getElementById('queryColumn').value = state.config.queryColumn || '测点';
  document.getElementById('interval').value = state.config.interval || '3600';
  document.getElementById('aggregationFunction').value = state.config.aggregationFunction || '';
  document.getElementById('dataFormat').value = state.config.dataFormat || 'latest';
  
  // 加载时间范围
  if (state.config.startTime) {
    document.getElementById('startTime').value = state.config.startTime;
  }
  if (state.config.endTime) {
    document.getElementById('endTime').value = state.config.endTime;
  }
  
  // 确保使用代理服务器
  if (!state.config.apiUrl || state.config.apiUrl.includes('digitaltwin.meta42.indc.vnet.com')) {
    state.config.apiUrl = 'http://localhost:3001';
    document.getElementById('apiUrl').value = 'http://localhost:3001';
  }
  
  log('配置已加载（使用默认值）', 'info');
}

// 保存配置
function saveConfig() {
  const sourceTableUrl = document.getElementById('sourceTableUrl').value;
  const targetTableUrl = document.getElementById('targetTableUrl').value;
  
  // 解析数据源表格 URL
  if (sourceTableUrl) {
    const parsed = parseTableUrl(sourceTableUrl);
    if (parsed && parsed.isValid) {
      state.config.sourceTableUrl = sourceTableUrl;
      state.config.sourceTableId = parsed.tableId;
      
      document.getElementById('sourceTableType').textContent = 
        parsed.type === 'base' ? 'Base 多维表格' : 'Wiki 多维表格';
      document.getElementById('sourceTableId').textContent = parsed.tableId || parsed.baseId || '-';
      
      log(`数据源表格 URL 已识别: ${parsed.type}`, 'success');
      
      // 尝试连接数据源表格
      connectToSourceTable(sourceTableUrl, parsed.tableId);
    } else {
      log('数据源表格 URL 格式不正确', 'error');
    }
  }
  
  // 解析目标表格 URL
  if (targetTableUrl) {
    const parsed = parseTableUrl(targetTableUrl);
    if (parsed && parsed.isValid) {
      state.config.targetTableUrl = targetTableUrl;
      state.config.targetTableId = parsed.tableId;
      
      document.getElementById('targetTableType').textContent = 
        parsed.type === 'base' ? 'Base 多维表格' : 'Wiki 多维表格';
      document.getElementById('targetTableId').textContent = parsed.tableId || parsed.baseId || '-';
      
      log(`目标表格 URL 已识别: ${parsed.type}`, 'success');
      
      // 尝试连接目标表格
      connectToTargetTable(targetTableUrl, parsed.tableId);
    } else {
      log('目标表格 URL 格式不正确', 'error');
    }
  }
  
  state.config.appId = document.getElementById('appId').value;
  state.config.appSecret = document.getElementById('appSecret').value;
  state.config.apiUrl = document.getElementById('apiUrl').value;
  state.config.authHeader = document.getElementById('authHeader').value;
  state.config.queryColumn = document.getElementById('queryColumn').value;
  state.config.startTime = document.getElementById('startTime').value;
  state.config.endTime = document.getElementById('endTime').value;
  state.config.interval = document.getElementById('interval').value;
  state.config.aggregationFunction = document.getElementById('aggregationFunction').value;
  state.config.dataFormat = document.getElementById('dataFormat').value;
  
  // 更新 API 配置
  if (state.api) {
    state.api.updateConfig({
      baseUrl: state.config.apiUrl,
      authHeader: state.config.authHeader
    });
  }
  
  localStorage.setItem('syncPluginConfig', JSON.stringify(state.config));
  log('配置已保存', 'success');
}

// 连接到数据源表格
async function connectToSourceTable(url, tableId) {
  try {
    log('正在连接数据源表格...', 'info');
    state.sourceTable = await getTableByUrl(url, tableId);
    log('✓ 已成功连接到数据源表格', 'success');
    
    // 显示数据源表格信息
    const tableName = await state.sourceTable.getName();
    log(`数据源表格名称: ${tableName}`, 'info');
  } catch (error) {
    log(`✗ 连接数据源表格失败: ${error.message}`, 'error');
    log('将使用当前表格作为数据源', 'warning');
    
    // 回退到当前表格
    state.sourceTable = await bitable.base.getActiveTable();
  }
}

// 连接到目标表格
async function connectToTargetTable(url, tableId) {
  try {
    log('正在连接目标表格...', 'info');
    state.targetTable = await getTableByUrl(url, tableId);
    log('✓ 已成功连接到目标表格', 'success');
    
    // 显示目标表格信息
    const tableName = await state.targetTable.getName();
    log(`目标表格名称: ${tableName}`, 'info');
  } catch (error) {
    log(`✗ 连接目标表格失败: ${error.message}`, 'error');
    log('将使用当前表格', 'warning');
    
    // 回退到当前表格
    state.targetTable = await bitable.base.getActiveTable();
  }
}

// 绑定事件
function bindEvents() {
  document.getElementById('saveConfig').addEventListener('click', saveConfig);
  document.getElementById('startSync').addEventListener('click', startSync);
  document.getElementById('pauseSync').addEventListener('click', pauseSync);
  document.getElementById('resumeSync').addEventListener('click', resumeSync);
  document.getElementById('stopSync').addEventListener('click', stopSync);
  document.getElementById('testSourceConnection').addEventListener('click', testSourceTableConnection);
  document.getElementById('testTargetConnection').addEventListener('click', testTargetTableConnection);
  
  // 数据源表格 URL 输入框实时解析
  document.getElementById('sourceTableUrl').addEventListener('input', (e) => {
    const url = e.target.value;
    if (url) {
      const parsed = parseTableUrl(url);
      if (parsed && parsed.isValid) {
        document.getElementById('sourceTableType').textContent = 
          parsed.type === 'base' ? 'Base 多维表格' : 'Wiki 多维表格';
        document.getElementById('sourceTableId').textContent = parsed.tableId || parsed.baseId || '-';
      } else {
        document.getElementById('sourceTableType').textContent = '无效链接';
        document.getElementById('sourceTableId').textContent = '-';
      }
    } else {
      document.getElementById('sourceTableType').textContent = '当前表格';
      document.getElementById('sourceTableId').textContent = '-';
    }
  });
  
  // 目标表格 URL 输入框实时解析
  document.getElementById('targetTableUrl').addEventListener('input', (e) => {
    const url = e.target.value;
    if (url) {
      const parsed = parseTableUrl(url);
      if (parsed && parsed.isValid) {
        document.getElementById('targetTableType').textContent = 
          parsed.type === 'base' ? 'Base 多维表格' : 'Wiki 多维表格';
        document.getElementById('targetTableId').textContent = parsed.tableId || parsed.baseId || '-';
      } else {
        document.getElementById('targetTableType').textContent = '无效链接';
        document.getElementById('targetTableId').textContent = '-';
      }
    } else {
      document.getElementById('targetTableType').textContent = '当前表格';
      document.getElementById('targetTableId').textContent = '-';
    }
  });
}

// 测试数据源表格连接
async function testSourceTableConnection() {
  const tableUrl = document.getElementById('sourceTableUrl').value;
  
  if (!tableUrl) {
    log('未配置数据源表格，将使用当前表格', 'info');
    return;
  }
  
  const parsed = parseTableUrl(tableUrl);
  if (!parsed || !parsed.isValid) {
    log('数据源表格 URL 格式不正确', 'error');
    return;
  }
  
  log('正在测试数据源表格连接...', 'info');
  
  try {
    const testTable = await getTableByUrl(tableUrl, parsed.tableId);
    const tableName = await testTable.getName();
    const fieldList = await testTable.getFieldMetaList();
    const recordCount = (await testTable.getRecordIdList()).length;
    
    log('✓ 连接成功！', 'success');
    log(`表格名称: ${tableName}`, 'info');
    log(`字段数量: ${fieldList.length}`, 'info');
    log(`记录数量: ${recordCount}`, 'info');
    
    // 检查是否有测点列
    const queryField = fieldList.find(f => f.name === state.config.queryColumn);
    if (queryField) {
      log(`✓ 找到"${state.config.queryColumn}"列`, 'success');
    } else {
      log(`⚠ 未找到"${state.config.queryColumn}"列，请确认列名配置`, 'warning');
    }
    
  } catch (error) {
    log(`✗ 连接失败: ${error.message}`, 'error');
    console.error('连接测试失败:', error);
  }
}

// 测试目标表格连接
async function testTargetTableConnection() {
  const tableUrl = document.getElementById('tableUrl').value;
  
  if (!tableUrl) {
    log('请先输入目标表格链接', 'error');
    return;
  }
  
  const parsed = parseTableUrl(tableUrl);
  if (!parsed || !parsed.isValid) {
    log('表格 URL 格式不正确', 'error');
    return;
  }
  
  log('正在测试连接...', 'info');
  
  try {
    const testTable = await getTableByUrl(tableUrl, parsed.tableId);
    const tableName = await testTable.getName();
    const fieldList = await testTable.getFieldMetaList();
    const recordCount = (await testTable.getRecordIdList()).length;
    
    log('✓ 连接成功！', 'success');
    log(`表格名称: ${tableName}`, 'info');
    log(`字段数量: ${fieldList.length}`, 'info');
    log(`记录数量: ${recordCount}`, 'info');
    
    // 检查是否有测点列
    const queryField = fieldList.find(f => f.name === state.config.queryColumn);
    if (queryField) {
      log(`✓ 找到"${state.config.queryColumn}"列`, 'success');
    } else {
      log(`⚠ 未找到"${state.config.queryColumn}"列，请确认列名配置`, 'warning');
    }
    
  } catch (error) {
    log(`✗ 连接失败: ${error.message}`, 'error');
    console.error('连接测试失败:', error);
  }
}

// 开始同步
async function startSync() {
  if (state.isSyncing) return;
  
  // 重置进度
  state.syncProgress = {
    currentIndex: 0,
    recordIdList: [],
    total: 0,
    successCount: 0,
    failCount: 0,
    isWriting: false,
    currentPointId: null
  };
  
  state.isSyncing = true;
  state.isPaused = false;
  document.getElementById('startSync').disabled = true;
  document.getElementById('pauseSync').disabled = false;
  document.getElementById('resumeSync').disabled = true;
  document.getElementById('stopSync').disabled = false;
  
  updateStatus('同步中...');
  log('开始数据拉取', 'info');
  
  try {
    await performSync();
  } catch (error) {
    log(`同步失败: ${error.message}`, 'error');
  } finally {
    if (!state.isPaused) {
      stopSync();
    }
  }
}

// 暂停同步
function pauseSync() {
  if (!state.isSyncing || state.isPaused) return;
  
  state.isPaused = true;
  document.getElementById('pauseSync').disabled = true;
  document.getElementById('resumeSync').disabled = false;
  
  updateStatus('已暂停');
  
  // 如果当前正在写入数据，需要回退索引以便继续时重新处理该测点
  if (state.syncProgress.isWriting && state.syncProgress.currentIndex > 0) {
    state.syncProgress.currentIndex--;
    log(`已暂停，当前测点 ${state.syncProgress.currentPointId} 未完全写入，继续时将重新处理`, 'warning');
    log(`当前进度: ${state.syncProgress.currentIndex}/${state.syncProgress.total}`, 'warning');
  } else {
    log(`已暂停，当前进度: ${state.syncProgress.currentIndex}/${state.syncProgress.total}`, 'warning');
  }
  
  log('可以更换目标表格后点击"继续"按钮', 'info');
  
  // 保存进度到 localStorage
  saveProgress();
}

// 继续同步
async function resumeSync() {
  if (!state.isPaused) return;
  
  state.isPaused = false;
  state.isSyncing = true;
  document.getElementById('pauseSync').disabled = false;
  document.getElementById('resumeSync').disabled = true;
  
  updateStatus('同步中...');
  log(`继续数据拉取，从第 ${state.syncProgress.currentIndex + 1} 条开始`, 'info');
  
  // 检查目标表格是否更换
  const currentTargetUrl = document.getElementById('targetTableUrl').value;
  if (currentTargetUrl !== state.config.targetTableUrl) {
    log('检测到目标表格已更换，正在重新连接...', 'info');
    
    // 保存新的目标表格配置
    saveConfig();
    
    // 重新连接目标表格
    const parsed = parseTableUrl(currentTargetUrl);
    if (parsed && parsed.isValid) {
      try {
        await connectToTargetTable(currentTargetUrl, parsed.tableId);
        log('✓ 已连接到新的目标表格', 'success');
        
        // 确保目标表格有必要的字段
        await ensureTargetTableFields();
      } catch (error) {
        log(`✗ 连接新目标表格失败: ${error.message}`, 'error');
        state.isPaused = true;
        state.isSyncing = false;
        document.getElementById('pauseSync').disabled = true;
        document.getElementById('resumeSync').disabled = false;
        return;
      }
    }
  }
  
  try {
    await performSync(true); // 传入 true 表示继续模式
  } catch (error) {
    log(`同步失败: ${error.message}`, 'error');
  } finally {
    if (!state.isPaused) {
      stopSync();
    }
  }
}

// 停止同步
function stopSync() {
  state.isSyncing = false;
  state.isPaused = false;
  document.getElementById('startSync').disabled = false;
  document.getElementById('pauseSync').disabled = true;
  document.getElementById('resumeSync').disabled = true;
  document.getElementById('stopSync').disabled = true;
  updateStatus('已停止');
  
  // 清除保存的进度
  clearProgress();
  log('同步已完全停止，进度已清除', 'info');
}

// 执行同步
async function performSync(isResume = false) {
  // 验证时间范围
  const startTime = new Date(state.config.startTime).getTime();
  const endTime = new Date(state.config.endTime).getTime();
  
  if (!startTime || !endTime) {
    throw new Error('请设置有效的时间范围');
  }
  
  if (startTime >= endTime) {
    throw new Error('开始时间必须早于结束时间');
  }
  
  if (!isResume) {
    log(`时间范围: ${state.config.startTime} 至 ${state.config.endTime}`, 'info');
    log(`时间间隔: ${state.config.interval} 秒`, 'info');
    log(`聚合函数: ${state.config.aggregationFunction || '无'}`, 'info');
  }
  
  // 显示数据源和目标表格信息
  const sourceName = await state.sourceTable.getName();
  const targetName = await state.targetTable.getName();
  log(`数据源表格: ${sourceName}`, 'info');
  log(`目标表格: ${targetName}`, 'info');
  
  // 从数据源表格获取所有字段
  const sourceFieldList = await state.sourceTable.getFieldMetaList();
  
  // 查找测点列
  const queryField = sourceFieldList.find(f => f.name === state.config.queryColumn);
  if (!queryField) {
    throw new Error(`在数据源表格中未找到列: ${state.config.queryColumn}`);
  }
  
  // 从数据源表格获取所有记录（如果不是继续模式）
  if (!isResume || state.syncProgress.recordIdList.length === 0) {
    state.syncProgress.recordIdList = await state.sourceTable.getRecordIdList();
    state.syncProgress.total = state.syncProgress.recordIdList.length;
    state.syncProgress.currentIndex = 0;
    state.syncProgress.successCount = 0;
    state.syncProgress.failCount = 0;
    
    log(`从数据源表格找到 ${state.syncProgress.total} 条记录`, 'info');
  } else {
    log(`继续处理剩余 ${state.syncProgress.total - state.syncProgress.currentIndex} 条记录`, 'info');
  }
  
  // 逐条处理（从当前进度开始）
  for (let i = state.syncProgress.currentIndex; i < state.syncProgress.recordIdList.length; i++) {
    // 检查是否暂停
    if (state.isPaused) {
      log('同步已暂停', 'warning');
      return;
    }
    
    if (!state.isSyncing) {
      log('同步已停止', 'warning');
      return;
    }
    
    const recordId = state.syncProgress.recordIdList[i];
    
    try {
      // 从数据源表格读取测点ID
      const cellValue = await state.sourceTable.getCellValue(queryField.id, recordId);
      console.log(`记录 ${i + 1} 原始单元格值:`, cellValue, '类型:', typeof cellValue);
      
      const pointId = extractQueryId(cellValue);
      console.log(`记录 ${i + 1} 提取的测点ID:`, pointId);
      
      if (!pointId) {
        log(`记录 ${i + 1}: 测点为空，跳过`, 'info');
        state.syncProgress.currentIndex = i + 1;
        state.syncProgress.isWriting = false;
        state.syncProgress.currentPointId = null;
        continue;
      }
      
      // 标记当前正在处理的测点
      state.syncProgress.currentPointId = pointId;
      state.syncProgress.isWriting = false;
      
      log(`记录 ${i + 1}/${state.syncProgress.total}: 正在查询 ${pointId}...`, 'info');
      
      // 查询历史数据（单个测点）
      const data = await fetchHistoryData([pointId], startTime, endTime);
      
      // 请求完成后延时 500ms
      await sleep(500);
      
      // 标记正在写入
      state.syncProgress.isWriting = true;
      
      // 写入数据到目标表格
      await writeHistoryData(pointId, data);
      
      // 写入完成，清除标记
      state.syncProgress.isWriting = false;
      state.syncProgress.currentPointId = null;
      
      // 写入完成后延时 200ms
      await sleep(200);
      
      state.syncProgress.successCount++;
      log(`记录 ${i + 1}/${state.syncProgress.total}: ${pointId} - 成功`, 'success');
    } catch (error) {
      state.syncProgress.failCount++;
      state.syncProgress.isWriting = false;
      state.syncProgress.currentPointId = null;
      log(`记录 ${i + 1}/${state.syncProgress.total}: 失败 - ${error.message}`, 'error');
      
      // 失败后也延时 200ms，避免连续失败导致请求过快
      await sleep(200);
    }
    
    // 更新当前索引
    state.syncProgress.currentIndex = i + 1;
    
    // 更新进度
    updateProgress(
      state.syncProgress.currentIndex, 
      state.syncProgress.total, 
      state.syncProgress.successCount, 
      state.syncProgress.failCount
    );
    
    // 定期保存进度（每10条保存一次）
    if (state.syncProgress.currentIndex % 10 === 0) {
      saveProgress();
    }
  }
  
  log(`同步完成！成功: ${state.syncProgress.successCount}, 失败: ${state.syncProgress.failCount}`, 'success');
}

// 提取查询ID
function extractQueryId(cellValue) {
  if (!cellValue) return null;
  
  let rawId = '';
  
  // 处理不同的数据类型
  if (typeof cellValue === 'string') {
    rawId = cellValue;
  } else if (typeof cellValue === 'number') {
    rawId = String(cellValue);
  } else if (Array.isArray(cellValue)) {
    // 飞书多维表格的富文本格式：数组中包含文本片段和链接
    // 例如：[{type: 'text', text: '29.125073'}, {type: 'url', text: '1.1.8.1.1.1.2', link: '...'}, {type: 'text', text: '.1'}]
    rawId = cellValue.map(item => {
      if (typeof item === 'string') {
        return item;
      } else if (typeof item === 'number') {
        return String(item);
      } else if (item && typeof item === 'object') {
        // 提取文本内容，忽略链接等其他属性
        return item.text || item.value || item.content || '';
      }
      return '';
    }).join('');
  } else if (typeof cellValue === 'object' && cellValue !== null) {
    // 处理单个对象格式
    rawId = cellValue.text || cellValue.value || cellValue.content || String(cellValue);
  } else {
    rawId = String(cellValue);
  }
  
  // 清理并返回
  const cleaned = DataTransformer.cleanDeviceId(rawId);
  console.log('extractQueryId - 输入:', JSON.stringify(cellValue), '输出:', cleaned);
  return cleaned;
}

// 请求历史数据
async function fetchHistoryData(pointList, startTime, endTime) {
  return await state.api.queryHistoryData({
    startTime,
    endTime,
    interval: state.config.interval,
    function: state.config.aggregationFunction,
    pointList
  });
}

// 写入历史数据到表格
async function writeHistoryData(pointId, apiData) {
  console.log('writeHistoryData - 开始写入:', { pointId, apiData });
  
  const parsed = DataTransformer.parseHistoryResponse(pointId, apiData);
  console.log('writeHistoryData - 解析结果:', parsed);
  
  if (!parsed.success) {
    throw new Error(parsed.error);
  }
  
  if (parsed.dataPoints.length === 0) {
    log(`测点 ${pointId} 没有数据`, 'warning');
    return;
  }
  
  // 确保目标表格有必要的字段
  const fields = await ensureTargetTableFields();
  
  // 验证字段对象
  if (!fields.pointField || !fields.timestampField || !fields.timeField || !fields.valueField) {
    throw new Error('字段对象无效，无法写入数据');
  }
  
  console.log('使用的字段ID:', {
    pointField: fields.pointField.id,
    timestampField: fields.timestampField.id,
    timeField: fields.timeField.id,
    valueField: fields.valueField.id
  });
  
  // 为每个数据点创建一条记录
  let createdCount = 0;
  let updatedCount = 0;
  let failedCount = 0;
  
  for (const dataPoint of parsed.dataPoints) {
    try {
      // 在目标表格中查找是否已有该记录（通过测点ID和时间戳匹配）
      const targetRecordIdList = await state.targetTable.getRecordIdList();
      let targetRecordId = null;
      
      // 尝试找到已有的记录
      for (const recId of targetRecordIdList) {
        try {
          const pointCellValue = await state.targetTable.getCellValue(fields.pointField.id, recId);
          const timestampCellValue = await state.targetTable.getCellValue(fields.timestampField.id, recId);
          
          const existingPointId = extractQueryId(pointCellValue);
          const existingTimestamp = String(timestampCellValue || '');
          
          if (existingPointId === pointId && existingTimestamp === String(dataPoint.timestamp)) {
            targetRecordId = recId;
            break;
          }
        } catch (error) {
          // 读取单元格失败，跳过该记录
          console.warn(`读取记录 ${recId} 失败:`, error);
          continue;
        }
      }
      
      // 如果没找到，创建新记录
      if (!targetRecordId) {
        try {
          targetRecordId = await state.targetTable.addRecord({
            fields: {
              [fields.pointField.id]: pointId,
              [fields.timestampField.id]: String(dataPoint.timestamp),
              [fields.timeField.id]: dataPoint.formattedTime,
              [fields.valueField.id]: String(dataPoint.value)
            }
          });
          createdCount++;
          console.log(`创建新记录: ${pointId} - ${dataPoint.formattedTime} - ${dataPoint.value}`);
        } catch (error) {
          failedCount++;
          console.error(`创建记录失败:`, error);
          log(`写入数据点失败: ${dataPoint.formattedTime} - ${error.message}`, 'error');
        }
      } else {
        // 更新已有记录
        try {
          await state.targetTable.setCellValue(fields.timeField.id, targetRecordId, dataPoint.formattedTime);
          await state.targetTable.setCellValue(fields.valueField.id, targetRecordId, String(dataPoint.value));
          updatedCount++;
          console.log(`更新记录: ${pointId} - ${dataPoint.formattedTime} - ${dataPoint.value}`);
        } catch (error) {
          failedCount++;
          console.error(`更新记录失败:`, error);
          log(`更新数据点失败: ${dataPoint.formattedTime} - ${error.message}`, 'error');
        }
      }
      
    } catch (error) {
      failedCount++;
      console.error(`处理数据点失败:`, error);
      log(`处理数据点失败: ${dataPoint.formattedTime} - ${error.message}`, 'error');
    }
  }
  
  if (failedCount > 0) {
    log(`测点 ${pointId}: 创建 ${createdCount} 条，更新 ${updatedCount} 条，失败 ${failedCount} 条`, 'warning');
  } else {
    log(`测点 ${pointId}: 创建 ${createdCount} 条，更新 ${updatedCount} 条记录`, 'success');
  }
  console.log('writeHistoryData - 写入完成');
}

// 确保目标表格有必要的字段
async function ensureTargetTableFields() {
  let fieldList = await state.targetTable.getFieldMetaList();
  let fieldsCreated = false;
  
  // 查找或创建"测点"字段
  let pointField = fieldList.find(f => f.name === '测点');
  if (!pointField) {
    const pointFieldId = await state.targetTable.addField({
      type: 1, // 文本类型
      name: '测点'
    });
    pointField = { id: pointFieldId, name: '测点' };
    log('创建字段: 测点', 'info');
    fieldsCreated = true;
  }
  
  // 查找或创建"时间戳"字段
  let timestampField = fieldList.find(f => f.name === '时间戳');
  if (!timestampField) {
    const timestampFieldId = await state.targetTable.addField({
      type: 1, // 文本类型
      name: '时间戳'
    });
    timestampField = { id: timestampFieldId, name: '时间戳' };
    log('创建字段: 时间戳', 'info');
    fieldsCreated = true;
  }
  
  // 查找或创建"时间"字段（格式化的时间）
  let timeField = fieldList.find(f => f.name === '时间');
  if (!timeField) {
    const timeFieldId = await state.targetTable.addField({
      type: 1, // 文本类型
      name: '时间'
    });
    timeField = { id: timeFieldId, name: '时间' };
    log('创建字段: 时间', 'info');
    fieldsCreated = true;
  }
  
  // 查找或创建"值"字段
  let valueField = fieldList.find(f => f.name === '值');
  if (!valueField) {
    const valueFieldId = await state.targetTable.addField({
      type: 1, // 文本类型
      name: '值'
    });
    valueField = { id: valueFieldId, name: '值' };
    log('创建字段: 值', 'info');
    fieldsCreated = true;
  }
  
  // 如果创建了新字段，重新获取字段列表以确保字段ID正确
  if (fieldsCreated) {
    log('重新获取字段列表以确保字段ID正确...', 'info');
    fieldList = await state.targetTable.getFieldMetaList();
    
    pointField = fieldList.find(f => f.name === '测点');
    timestampField = fieldList.find(f => f.name === '时间戳');
    timeField = fieldList.find(f => f.name === '时间');
    valueField = fieldList.find(f => f.name === '值');
    
    // 验证所有字段都找到了
    if (!pointField || !timestampField || !timeField || !valueField) {
      throw new Error('字段创建后无法找到，请检查表格权限');
    }
    
    log('✓ 所有字段已就绪', 'success');
  }
  
  return {
    pointField,
    timestampField,
    timeField,
    valueField
  };
}

// 保存进度到 localStorage
function saveProgress() {
  const progress = {
    currentIndex: state.syncProgress.currentIndex,
    recordIdList: state.syncProgress.recordIdList,
    total: state.syncProgress.total,
    successCount: state.syncProgress.successCount,
    failCount: state.syncProgress.failCount,
    isWriting: state.syncProgress.isWriting,
    currentPointId: state.syncProgress.currentPointId,
    timestamp: Date.now()
  };
  localStorage.setItem('syncProgress', JSON.stringify(progress));
  console.log('进度已保存:', progress);
}

// 加载进度
function loadProgress() {
  const saved = localStorage.getItem('syncProgress');
  if (saved) {
    try {
      const progress = JSON.parse(saved);
      
      // 检查进度是否过期（超过24小时）
      const age = Date.now() - progress.timestamp;
      if (age > 24 * 3600 * 1000) {
        log('保存的进度已过期，已清除', 'info');
        clearProgress();
        return false;
      }
      
      state.syncProgress = {
        currentIndex: progress.currentIndex || 0,
        recordIdList: progress.recordIdList || [],
        total: progress.total || 0,
        successCount: progress.successCount || 0,
        failCount: progress.failCount || 0,
        isWriting: false,  // 加载时重置写入状态
        currentPointId: progress.currentPointId || null
      };
      
      if (progress.currentPointId) {
        log(`发现未完成的同步任务，进度: ${state.syncProgress.currentIndex}/${state.syncProgress.total}`, 'info');
        log(`上次暂停时正在处理测点: ${progress.currentPointId}，继续时将重新处理`, 'info');
      } else {
        log(`发现未完成的同步任务，进度: ${state.syncProgress.currentIndex}/${state.syncProgress.total}`, 'info');
      }
      return true;
    } catch (error) {
      console.error('加载进度失败:', error);
      clearProgress();
      return false;
    }
  }
  return false;
}

// 清除进度
function clearProgress() {
  localStorage.removeItem('syncProgress');
  state.syncProgress = {
    currentIndex: 0,
    recordIdList: [],
    total: 0,
    successCount: 0,
    failCount: 0,
    isWriting: false,
    currentPointId: null
  };
}

// 工具函数
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function updateStatus(status) {
  document.getElementById('status').textContent = status;
}

function updateProgress(current, total, success, fail) {
  document.getElementById('progress').textContent = `${current}/${total}`;
  document.getElementById('successCount').textContent = success;
  document.getElementById('failCount').textContent = fail;
}

function log(message, type = 'info') {
  const logContainer = document.getElementById('logContainer');
  const logItem = document.createElement('div');
  logItem.className = `log-item ${type}`;
  const time = new Date().toLocaleTimeString();
  logItem.textContent = `[${time}] ${message}`;
  logContainer.appendChild(logItem);
  logContainer.scrollTop = logContainer.scrollHeight;
}

// 启动
init();
