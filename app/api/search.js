const request = require('./request');

const contextSearch = async (keywords) => {
  try {
    const res = await request({
      url: '/Search/ContextSearch',
      method: 'GET',
      params: {
        Keywords: keywords
      }
    });
    return res;
  } catch (error) {
    console.log('[Request Error] Failed to get preContext', error);
    throw error;
  }
};

const keySearch = async (keywords) => {
  try {
    const res = await request({
      url: '/Search/KeySearch',
      method: 'GET',
      params: {
        Keywords: keywords
      }
    });
    return res;
  } catch (error) {
    console.log('[Request Error] Failed to get keySearch results', error);
    throw error;
  }
};

module.exports = {
  contextSearch,
  keySearch
};