const CardFactory = require("botbuilder").CardFactory;

const createSearchCard = (
  query = "",
  isAISearch = true,
  selectedTypes = "1,2,3,4,5"
) => {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.0",
    style: "default",
    width: "stretch",
    body: [
      {
        type: "TextBlock",
        text: "AI search",
        weight: "bolder",
        size: "medium",
      },
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "Input.Text",
                id: "searchQuery",
                placeholder: "Enter your question here ...",
                value: query,
                isRequired: true,
                errorMessage: "Please enter a question.",
              },
            ],
          },
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "ActionSet",
                actions: [
                  {
                    type: "Action.Submit",
                    title: "🔍",
                    data: { action: "aiSearch" },
                  },
                ],
              },
            ],
          },
        ],
      },
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "Input.ChoiceSet",
                id: "searchTypes",
                style: "compact",
                isMultiSelect: true,
                value: selectedTypes,
                placeholder: "Select search types",
                choices: [
                  {
                    title: "Account",
                    value: "1",
                  },
                  {
                    title: "Contact",
                    value: "2",
                  },
                  {
                    title: "Fund",
                    value: "3",
                  },
                  {
                    title: "Activity",
                    value: "4",
                  },
                  {
                    title: "Document",
                    value: "5",
                  },
                ],
              },
            ],
          },
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "Input.Toggle",
                id: "searchMode",
                title: "AI",
                valueOn: "true",
                valueOff: "false",
                value: isAISearch.toString(),
                wrap: false,
                style: "positive",
              },
            ],
          },
          {
            type: "Column",
            width: "53px",
            items: [
              {
                type: "Input.Number",
                id: "maxResultCount",
                placeholder: "5",
                max: 10,
                min: 1,
                value: 10,
              },
            ],
          },
        ],
      },
    ],
  });
};

module.exports = {
  createSearchCard,
};
