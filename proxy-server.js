/**
 * API 代理服务器
 * 用于解决 CORS 跨域问题
 */

const http = require('http');
const https = require('https');
const url = require('url');

const PORT = 3001;
const TARGET_API = 'https://digitaltwin.meta42.indc.vnet.com/openapi/tsdb/point_data/v2/search';
const AUTH_HEADER = 'Basic dGVjaG5pcXVlX2NlbnRlcjoyMVZpYW5ldEBWbmV0LmNvbQ==';

const server = http.createServer((req, res) => {
  // 设置 CORS 头
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  // 处理 OPTIONS 预检请求
  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // 只处理 POST 请求
  if (req.method !== 'POST') {
    res.writeHead(405, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Method not allowed' }));
    return;
  }

  // 收集请求数据
  let body = '';
  req.on('data', chunk => {
    body += chunk.toString();
  });

  req.on('end', () => {
    console.log(`[${new Date().toISOString()}] 收到请求`);
    console.log('请求数据:', body);

    // 解析目标 URL
    const targetUrl = url.parse(TARGET_API);

    // 配置 HTTPS 请求
    const options = {
      hostname: targetUrl.hostname,
      port: 443,
      path: targetUrl.path,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': AUTH_HEADER,
        'Content-Length': Buffer.byteLength(body)
      }
    };

    // 发送请求到目标 API
    const proxyReq = https.request(options, (proxyRes) => {
      console.log(`目标 API 响应: ${proxyRes.statusCode}`);

      let responseData = '';
      proxyRes.on('data', chunk => {
        responseData += chunk;
      });

      proxyRes.on('end', () => {
        console.log('响应数据:', responseData);

        // 返回响应给客户端
        res.writeHead(proxyRes.statusCode, {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        });
        res.end(responseData);
      });
    });

    proxyReq.on('error', (error) => {
      console.error('代理请求失败:', error.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: '代理请求失败',
        message: error.message
      }));
    });

    // 发送请求数据
    proxyReq.write(body);
    proxyReq.end();
  });
});

server.listen(PORT, () => {
  console.log('='.repeat(60));
  console.log('🚀 API 代理服务器已启动（历史数据查询）');
  console.log('='.repeat(60));
  console.log(`监听端口: http://localhost:${PORT}`);
  console.log(`目标 API: ${TARGET_API}`);
  console.log('');
  console.log('使用方法:');
  console.log(`1. 在插件的"高级配置"中修改 API 地址为: http://localhost:${PORT}`);
  console.log('2. 配置时间范围和查询参数');
  console.log('3. 保存配置并开始数据拉取');
  console.log('');
  console.log('按 Ctrl+C 停止服务器');
  console.log('='.repeat(60));
});

// 优雅关闭
process.on('SIGINT', () => {
  console.log('\n正在关闭服务器...');
  server.close(() => {
    console.log('服务器已关闭');
    process.exit(0);
  });
});
