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
  const token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJodHRwOi8vc2NoZW1hcy54bWxzb2FwLm9yZy93cy8yMDA1LzA1L2lkZW50aXR5L2NsYWltcy9uYW1laWRlbnRpZmllciI6IjkyMzY0Yjc1LWNhNjYtNDc4NC04MTlmLWU5ODRkM2ZjYThhYyIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiJqb2hubnkgd2FuZyIsIm5iZiI6MTczNjkyMjk2MywiZXhwIjoxNzM3NTI3NzYzLCJpc3MiOiJodHRwOi8vbG9jYWxob3N0OjUwMDAiLCJhdWQiOiJodHRwOi8vbG9jYWxob3N0OjUwMDAifQ.arpbXMuknid6ht6V-3qWGrewrwBwla7G5fFmdeftuqs";  // 替换为实际的 token
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