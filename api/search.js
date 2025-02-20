const request = require('./request');

const contextSearch = async (keywords, modulesFilterStr = 'TargetTypes=1&TargetTypes=2&TargetTypes=3&TargetTypes=4&TargetTypes=5', count = 10) => {
  try {
    const startTime = Date.now();
    const res = await request({
      url: `/Search/ContextSearch?${modulesFilterStr}&count=${count}`,
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

const keySearch = async (keywords, modulesFilterStr = 'TargetTypes=1&TargetTypes=2&TargetTypes=3', count = 10) => {
  try {
    const res = await request({
      url: `/Search/KeySearch?${modulesFilterStr}&count=${count}`,
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

const queryTags = async (type) => {
  const response = await request({
    url: `/menus/getMenus/${type}`,
    method: "GET",
  });
  return await response;
};

const queryAccountList = async (data) => {
  try {
    const res = await request({
      url: "/Accounts/GetPagingAccount",
      method: "POST",
      data,
    });
    return res;
  } catch (error) {
    console.log("[Request Error] Failed to get account list", error);
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

const queryFundList = async (data) => {
  try {
    const res = await request({
      url: "/Funds/GetPagingFund",
      method: "POST",
      data,
    });
    return res;
  } catch (error) {
    console.log("[Request Error] Failed to get fund list", error);
    throw error;
  }
};

const queryActivityList = async (data) => {
  try {
    const res = await request({
      url: "/Activities/GetPagingActivity",
      method: "POST",
      data,
    });
    return res;
  } catch (error) {
    console.log("[Request Error] Failed to get activity list", error);
    throw error;
  }
};

const queryDocumentList = async (data) => {
  try {
    const res = await request({
      url: "/Documents/GetPagingDocuments",
      method: "POST",
      data,
    });
    return res;
  } catch (error) {
    console.log("[Request Error] Failed to get document list", error);
    throw error;
  }
};

module.exports = {
  contextSearch,
  keySearch,
  queryTags,
  queryContactList,
  queryFundList,
  queryAccountList,
  queryActivityList,
  queryDocumentList
};