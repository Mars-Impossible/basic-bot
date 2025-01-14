const request = require('./request');

const contextSearch = async (keywords, modulesFilterStr = 'TargetTypes=1&TargetTypes=2&TargetTypes=3&TargetTypes=4&TargetTypes=5') => {
  try {
    const startTime = Date.now();
    const res = await request({
      url: `/Search/ContextSearch?${modulesFilterStr}`,
      method: 'GET',
      params: {
        Keywords: keywords
      }
    });
    const endTime = Date.now();
    console.log(`[ContextSearch] Request time: ${endTime - startTime}ms`);
    return res;
  } catch (error) {
    console.log('[Request Error] Failed to get preContext', error);
    throw error;
  }
};

const keySearch = async (keywords, modulesFilterStr = 'TargetTypes=1&TargetTypes=2&TargetTypes=3') => {
  try {
    const res = await request({
      url: `/Search/KeySearch?${modulesFilterStr}`,
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

const queryContactList = async (data) => {
  try {
    const res = await request({
      url: "/Contacts/GetPagingContact",
      method: "POST",
      data,
    });
    return res;
  } catch (error) {
    console.log("[Request Error] Failed to get contact list", error);
    throw error;
  }
};

module.exports = {
  contextSearch,
  keySearch,
  queryContactList,
};