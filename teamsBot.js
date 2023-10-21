const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const workLocationCard = require("./adaptiveCards/homeOrOffice.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");
class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.officeCountobj = 0;
    this.homeCountobj = 0;
    require('dotenv').config();
    const apiKey = process.env.API_KEY;

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declare(workLocationCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "office": {
          this.officeCountobj += 1;
          break;
        }
        case "home": {
          this.homeCountobj += 1;
          break;
        }
        case "show": {
          if (this.officeCountobj == 0 && this.homeCountobj == 0) {
            await context.sendActivity('No answers yet');
          } else {
            await context.sendActivity(`At the office: ${this.officeCountobj} person(s), at home: ${this.homeCountobj} person(s).`)
          }
          break;
        }
        case "weather": {
          try {
            const apiUrl = `https://api.openweathermap.org/data/2.5/weather?q=Espoo,fi&units=metric&APPID=${apiKey}`

            const response = await fetch(apiUrl);
            if (response.ok) {
              const weatherData = await response.json();

              const temperature = weatherData.main.temp
              const description = weatherData.weather[0].description

              const message = (`Weather in Espoo: ${temperature}°C, ${description}`)
              await context.sendActivity(message);
            }
            else {
              await context.sendActivity('Error fetching weather data');
            }
          }
          catch (error) {
            console.error(error);
          }
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;