const aiChatConfig = require('../store/aiChatConfig');

const createContactCard = (contact,detailUrl) => ({
  composeExtension: {
    type: 'result',
    attachmentLayout: 'list',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.0',
        body: [
          {
            type: 'TextBlock',
            text: contact.fullName || 'Unknown Name',
            weight: 'bolder',
            size: 'medium'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Email:', value: contact.email || 'N/A' },
              { title: 'Phone:', value: contact.mainPhone || contact.phone || 'N/A' },
              { title: 'Mobile:', value: contact.mainMobile || contact.mobile || 'N/A' },
              { title: 'Company:', value: contact.relatedAccount?.name || 'N/A' },
              { title: 'Address:', value: contact.mainAddress?.fullAddress || 'N/A' },
              { title: 'Overview:', value: contact.overview || 'N/A' }
            ]
          }
        ],
        actions: [
          {
            type: 'Action.OpenUrl',
            title: 'View Details',
            url: detailUrl
          }
        ]
      }
    }]
  }
});

const createAccountCard = (account,detailUrl) => ({
  composeExtension: {
    type: 'result',
    attachmentLayout: 'list',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.0',
        body: [
          {
            type: 'TextBlock',
            text: account.name || 'Unknown Account',
            weight: 'bolder',
            size: 'medium'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Email:', value: account.email || 'N/A' },
              { title: 'Phone:', value: account.mainPhone || account.phone || 'N/A' },
              { title: 'Domain:', value: account.domain || 'N/A' },
              { title: 'Address:', value: account.fullAddress || 'N/A' },
              { title: 'Overview:', value: account.overview || 'N/A' }
            ]
          }
        ],
        actions: [
          {
            type: 'Action.OpenUrl',
            title: 'View Details',
            url: detailUrl
          }
        ]
      }
    }]
  }
});

const createFundCard = (fund,detailUrl) => ({
  composeExtension: {
    type: 'result',
    attachmentLayout: 'list',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.0',
        body: [
          {
            type: 'TextBlock',
            text: fund.name || 'Unknown Fund',
            weight: 'bolder',
            size: 'medium'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Style:', value: fund.fund_Style || 'N/A' },
              { title: 'Type:', value: fund.sec_Type || 'N/A' },
              { title: 'Currency:', value: fund.currencyCode || 'N/A' },
              { title: 'Annual Return:', value: fund.annualReturn?.toString() + '%' || 'N/A' },
              { title: 'Annual Volatility:', value: fund.annualVolatility?.toString() + '%' || 'N/A' },
              { title: 'Overview:', value: fund.overview || 'N/A' }
            ]
          }
        ],
        actions: [
          {
            type: 'Action.OpenUrl',
            title: 'View Details',
            url: detailUrl
          }
        ]
      }
    }]
  }
});

const createActivityCard = (activity,detailUrl) => ({
  composeExtension: {
    type: 'result',
    attachmentLayout: 'list',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.0',
        body: [
          {
            type: 'TextBlock',
            text: activity.title || 'Unknown Activity',
            weight: 'bolder',
            size: 'medium'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'From:', value: activity.fromAddress || 'N/A' },
              { title: 'To:', value: activity.toAddress || 'N/A' },
              { title: 'CC:', value: activity.ccAddress || 'N/A' },
              { title: 'Contact:', value: activity.relatedContactItem?.name || 'N/A' },
              { title: 'Description:', value: activity.description || 'N/A' }
            ]
          }
        ],
        actions: [
          {
            type: 'Action.OpenUrl',
            title: 'View Details',
            url: detailUrl
          }
        ]
      }
    }]
  }
});

const createDocumentCard = (document,detailUrl) => ({
  composeExtension: {
    type: 'result',
    attachmentLayout: 'list',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.0',
        body: [
          {
            type: 'TextBlock',
            text: document.fileName || 'Unknown Document',
            weight: 'bolder',
            size: 'medium'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'From:', value: document.fromAddress || 'N/A' },
              { title: 'To:', value: document.toAddress || 'N/A' },
              { title: 'Description:', value: document.description || 'N/A' }
            ]
          }
        ],
        actions: [
          {
            type: 'Action.OpenUrl',
            title: 'View Details',
            url: detailUrl
          }
        ]
      }
    }]
  }
});

const createErrorCard = (message) => ({
  composeExtension: {
    type: 'result',
    attachmentLayout: 'list',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.0',
        body: [{
          type: 'TextBlock',
          text: message,
          color: 'attention'
        }]
      }
    }]
  }
});


const buildDetailUrl = (obj) => {
  const targetTypeObj = aiChatConfig.targetTypes.find(t => t.id === obj.targetType);
  if (!targetTypeObj) return 'https://newchat.arencore.me/';

  const baseModule = targetTypeObj.name.toLowerCase();
  
  let url = `https://newchat.arencore.me/${baseModule}?keywords=${encodeURIComponent(obj.relatedId)}`;
  url += `&module=${encodeURIComponent(baseModule)}`;
  
  if (obj.targetType !== 1) {
    url += `&tagMenuId=${encodeURIComponent(obj.tagMenuId)}`;
  }
  return url;
}

module.exports = {
  createContactCard,
  createAccountCard,
  createFundCard,
  createActivityCard,
  createDocumentCard,
  createErrorCard,
  buildDetailUrl
};
