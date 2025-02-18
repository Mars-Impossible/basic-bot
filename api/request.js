const axios = require('axios');

// TODO: set env baseURL
const request = axios.create({
  // baseURL: process.env.NEXT_PUBLIC_APP_API_BASE_URL,
  baseURL: 'https://marsai.arencore.me/api',
  timeout: 60000,
});

//TODO: 配置SSO token
request.interceptors.request.use((config) => {
  // 添加固定的 token
  const token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJodHRwOi8vc2NoZW1hcy54bWxzb2FwLm9yZy93cy8yMDA1LzA1L2lkZW50aXR5L2NsYWltcy9uYW1laWRlbnRpZmllciI6IjVmODAxN2ZhLTViNDYtNGY1Ny1hMzFjLTA1NjkxMTg1NzE2YiIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiJ4aWFvaGUiLCJuYmYiOjE3MzkxODQ5NzYsImV4cCI6MTczOTc4OTc3NiwiaXNzIjoiaHR0cDovL2xvY2FsaG9zdDo1MDAwIiwiYXVkIjoiaHR0cDovL2xvY2FsaG9zdDo1MDAwIn0.V5c3CtSufbBBnc0ht8RLNEzgYx6Psm3mzWwvFRdGnH0";  // 替换为实际的 token
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