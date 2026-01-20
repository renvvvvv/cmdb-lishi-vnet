/**
 * API 请求封装模块
 * 支持内网和外网地址切换，Basic Auth 认证
 */

export class DataAPI {
  constructor(config = {}) {
    this.config = {
      baseUrl: config.baseUrl || 'http://localhost:3001',
      authHeader: config.authHeader || 'Basic dGVjaG5pcXVlX2NlbnRlcjoyMVZpYW5ldEBWbmV0LmNvbQ==',
      timeout: config.timeout || 30000
    };
  }

  /**
   * 更新配置
   */
  updateConfig(config) {
    this.config = { ...this.config, ...config };
  }

  /**
   * 获取 Authorization Header
   */
  getAuthHeader() {
    return this.config.authHeader;
  }

  /**
   * 通用请求方法
   */
  async request(endpoint, options = {}) {
    const url = `${this.config.baseUrl}${endpoint}`;
    const headers = {
      'Content-Type': 'application/json',
      'Authorization': this.getAuthHeader(),
      ...options.headers
    };

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), this.config.timeout);

    try {
      console.log('API Request:', {
        url,
        method: options.method || 'GET',
        headers,
        body: options.body
      });

      const response = await fetch(url, {
        ...options,
        headers,
        signal: controller.signal,
        mode: 'cors',
        credentials: 'omit'
      });

      clearTimeout(timeoutId);

      console.log('API Response:', {
        status: response.status,
        statusText: response.statusText,
        ok: response.ok
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error('API Error Response:', errorText);
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const result = await response.json();
      console.log('API Response Data:', result);
      
      // 处理包装的响应格式
      if (result && typeof result === 'object') {
        // 如果有 data 字段，返回 data
        if (result.data) {
          return result.data;
        }
        // 否则返回整个结果
        return result;
      }
      
      return result;
    } catch (error) {
      clearTimeout(timeoutId);
      
      console.error('API Request Failed:', {
        error: error.message,
        name: error.name,
        url
      });
      
      if (error.name === 'AbortError') {
        throw new Error('请求超时');
      }
      
      if (error.message === 'Failed to fetch') {
        throw new Error('网络请求失败，可能是 CORS 跨域问题或网络不通。请检查：1) API 地址是否正确 2) 网络是否可访问 3) 服务器是否支持跨域请求');
      }
      
      throw error;
    }
  }

  /**
   * 查询历史数据
   * @param {Object} params - 查询参数
   * @param {number} params.startTime - 开始时间（毫秒时间戳）
   * @param {number} params.endTime - 结束时间（毫秒时间戳）
   * @param {string} params.interval - 时间间隔（秒）
   * @param {string} params.function - 聚合函数
   * @param {string[]} params.pointList - 测点列表
   * @returns {Promise<Object>} 历史数据
   */
  async queryHistoryData(params) {
    const { startTime, endTime, interval, function: func, pointList } = params;
    
    if (!pointList || pointList.length === 0) {
      throw new Error('测点列表不能为空');
    }

    const requestBody = {
      startTime,
      endTime,
      interval: String(interval),
      function: func || '',
      pointList
    };

    return await this.request('/tsdb/point_data/v2/search', {
      method: 'POST',
      body: JSON.stringify(requestBody)
    });
  }

  /**
   * 批量查询历史数据（分批处理）
   * @param {string[]} pointList - 测点列表
   * @param {Object} timeParams - 时间参数
   * @param {number} batchSize - 每批数量
   * @param {Function} onProgress - 进度回调
   * @returns {Promise<Object>} 合并后的结果
   */
  async batchQueryHistory(pointList, timeParams, batchSize = 50, onProgress = null) {
    const results = {};
    const batches = [];

    // 分批
    for (let i = 0; i < pointList.length; i += batchSize) {
      batches.push(pointList.slice(i, i + batchSize));
    }

    // 逐批查询
    for (let i = 0; i < batches.length; i++) {
      const batch = batches[i];
      const batchResult = await this.queryHistoryData({
        ...timeParams,
        pointList: batch
      });
      Object.assign(results, batchResult);

      if (onProgress) {
        onProgress({
          current: i + 1,
          total: batches.length,
          processed: (i + 1) * batchSize,
          totalItems: pointList.length
        });
      }

      // 批次间延时
      if (i < batches.length - 1) {
        await this.sleep(100);
      }
    }

    return results;
  }

  /**
   * 测试连接
   */
  async testConnection() {
    try {
      const now = Date.now();
      await this.queryHistoryData({
        startTime: now - 3600000,
        endTime: now,
        interval: '3600',
        function: '',
        pointList: ['test']
      });
      return { success: true, message: '连接成功' };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * 延时工具
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

/**
 * 数据转换工具
 */
export class DataTransformer {
  /**
   * 转换时间戳为可读格式
   */
  static formatTimestamp(timestamp) {
    if (!timestamp) return '';
    const date = new Date(parseInt(timestamp) * 1000);
    return date.toLocaleString('zh-CN', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
  }

  /**
   * 验证设备ID格式
   */
  static validateDeviceId(id) {
    if (!id || typeof id !== 'string') return false;
    // 格式：数字.数字.数字...
    return /^[\d.]+$/.test(id.trim());
  }

  /**
   * 清理设备ID
   */
  static cleanDeviceId(id) {
    if (!id) return null;
    return String(id).trim();
  }

  /**
   * 解析历史数据 API 响应
   */
  static parseHistoryResponse(pointId, responseData) {
    const data = responseData[pointId];
    
    if (!data || !Array.isArray(data) || data.length === 0) {
      return {
        success: false,
        dataPoints: [],
        count: 0,
        error: '未返回数据'
      };
    }

    // 解析数据点 - 支持两种格式
    const dataPoints = data.map(point => {
      // 格式1: 对象格式 {timestamp: xxx, value: xxx}
      if (point && typeof point === 'object' && point.timestamp !== undefined) {
        return {
          timestamp: point.timestamp,
          value: point.value,
          formattedTime: this.formatTimestamp(Math.floor(point.timestamp / 1000))
        };
      }
      // 格式2: 数组格式 [timestamp, value]
      else if (Array.isArray(point) && point.length >= 2) {
        return {
          timestamp: point[0],
          value: point[1],
          formattedTime: this.formatTimestamp(Math.floor(point[0] / 1000))
        };
      }
      // 其他格式
      return null;
    }).filter(p => p !== null);

    if (dataPoints.length === 0) {
      return {
        success: false,
        dataPoints: [],
        count: 0,
        error: '数据格式错误'
      };
    }

    return {
      success: true,
      dataPoints,
      count: dataPoints.length,
      latestValue: dataPoints[dataPoints.length - 1]?.value,
      latestTime: dataPoints[dataPoints.length - 1]?.formattedTime,
      error: null
    };
  }
  
  /**
   * 格式化历史数据为字符串（用于写入单元格）
   */
  static formatHistoryDataForCell(dataPoints, format = 'latest') {
    if (!dataPoints || dataPoints.length === 0) {
      return '';
    }
    
    switch (format) {
      case 'latest':
        // 只返回最新值
        const latest = dataPoints[dataPoints.length - 1];
        return `${latest.value} (${latest.formattedTime})`;
      
      case 'all':
        // 返回所有值
        return dataPoints.map(p => `${p.value} (${p.formattedTime})`).join('\n');
      
      case 'count':
        // 返回数据点数量
        return `${dataPoints.length} 个数据点`;
      
      default:
        return String(dataPoints[dataPoints.length - 1]?.value || '');
    }
  }
}

/**
 * 预设配置
 */
export const API_PRESETS = {
  internal: {
    name: '内网地址',
    baseUrl: 'http://gateway.meta42.indc.vnet.com/openapi',
    requiresHost: true,
    hostConfig: '192.168.233.91 gateway.meta42.indc.vnet.com'
  },
  external: {
    name: '外网地址',
    baseUrl: 'https://digitaltwin.meta42.indc.vnet.com/openapi',
    requiresHost: false
  },
  proxy: {
    name: '代理服务器',
    baseUrl: 'http://localhost:3001',
    requiresHost: false
  }
};

/**
 * 聚合函数选项
 */
export const AGGREGATION_FUNCTIONS = {
  none: { name: '无', value: '' },
  avg: { name: '平均值', value: 'avg' },
  sum: { name: '求和', value: 'sum' },
  max: { name: '最大值', value: 'max' },
  min: { name: '最小值', value: 'min' },
  count: { name: '计数', value: 'count' }
};
