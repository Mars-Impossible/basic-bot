const axios = require('axios');

// 创建 axios 实例
const request = axios.create({
  // baseURL: process.env.NEXT_PUBLIC_APP_API_BASE_URL,
  baseURL: 'https://marsai.arencore.me/api',
  timeout: 60000,
});

// 请求拦截器
request.interceptors.request.use((config) => {
  // 添加固定的 token
  const token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJodHRwOi8vc2NoZW1hcy54bWxzb2FwLm9yZy93cy8yMDA1LzA1L2lkZW50aXR5L2NsYWltcy9uYW1laWRlbnRpZmllciI6ImU0OThiYTI1LWE3NGMtNDk4Ni1iMDFjLTY0Y2EzMDUyOGZkMyIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiJqb2hubnkgd2FuZyIsIm5iZiI6MTczNjIxNDE4MSwiZXhwIjoxNzM2ODE4OTgxLCJpc3MiOiJodHRwOi8vbG9jYWxob3N0OjUwMDAiLCJhdWQiOiJodHRwOi8vbG9jYWxob3N0OjUwMDAifQ.ol8i91Wth3guqiQ8HeQMKW2bBZ4YqNGk0QyFM9ztrHI";  // 替换为实际的 token
  config.headers["Authorization"] = "Bearer " + token;
  return config;
});

// 响应拦截器
request.interceptors.response.use(
  (response) => response.data,
  (error) => {
    console.error('[Request Error]', error);
    return Promise.reject(error);
  }
);

module.exports = request;