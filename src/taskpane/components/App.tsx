import * as React from "react";
import axios from "axios";
import Progress from "./Progress";
import { DirectLine } from 'botframework-directlinejs';
import ReactWebChat from 'botframework-webchat';
import { ConfigKeys, Status } from '../models/Config';
import { Event, QnAMakerEndpoint } from '../models/Event';
import { QnADTO, Source } from "../models/QnAMaker";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
  store: any;
}

export interface AppState {
  webChatToken: string;
  debugstring: string[];
}

export default class App extends React.Component<AppProps, AppState> {

  tempUserId = 'TempUserId';
  toDispatch = [];
  buttonSyncConfig = 'SyncConfig';
  buttonSyncQAs = 'SyncQAs';
  buttonSyncQA = 'SyncQA';
  buttonCreateQA = "CreateQA";
  newQAName = "New QA";

  constructor(props: AppProps, context: any) {
    super(props, context);
    this.state = {
      webChatToken: '',
      debugstring: []
    };
  }

  addDebug(error: any) {
    this.setState({debugstring: this.state.debugstring.concat(String(error))});
  }

  clearDebug() {
    this.setState({debugstring: []});
  }

  clickClearDebug = async () => {
    this.clearDebug();
  };

  pushEvent(name: string, value: any) {
    this.toDispatch.push({
      type: 'WEB_CHAT/SEND_EVENT',
      payload: {name: name, value: value}
    });

    // this.addDebug(`${name}:${value}`);
  }

  disableButton(name: string, disable: boolean) {
    (document.getElementById(name) as HTMLButtonElement).disabled = disable;
  }

  async getConfig(context: Excel.RequestContext) {
    const configSheet = context.workbook.worksheets.getFirst();
    const configRange = configSheet.getUsedRange();
    configRange.load("values");

    await context.sync();

    let result = new Map();
    for (let i = 0;i < configRange.values.length;++i) {
      let element = configRange.values[i];
      if (element.length < 2) continue;
      result.set(String(element[0]).toLowerCase(), element[1]);
    }
    return result;
  }

  getTokenId = async () => {
    try {
      // TODO why?
      if (this.state.webChatToken != '') {
        return;
      }

      await Excel.run(async context => {
        let token = String((await this.getConfig(context)).get(ConfigKeys.WebChatToken));
        this.setState({ webChatToken: token });
        this.addDebug(token);
      });
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickSyncConfig = async() => {
    try {
      this.disableButton(this.buttonSyncConfig, true);

      await Excel.run(async context => {
        let config = await this.getConfig(context);

        let pushed = false;
        config.forEach((value, key) => {
          switch (key) {
            case ConfigKeys.ResultNumber:
              this.pushEvent(Event.SetResultNumber, Number(value));
              pushed = true;
              break;
            case ConfigKeys.NoResultResponse:
              this.pushEvent(Event.SetNoResultResponse, String(value));
              pushed = true;
              break;
            case ConfigKeys.MinScore:
              this.pushEvent(Event.SetMinScore, Number(value));
              pushed = true;
              break;
            case ConfigKeys.Debug:
              this.pushEvent(Event.SetDebug, Boolean(value));
              pushed = true;
              break;
          }
        });
        if (!pushed) {
          this.disableButton(this.buttonSyncConfig, false);
        }
      });

      // TODO why
      setTimeout(() => {this.clickDoSync()}, 1000);
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickSyncQAs = async () => {
    this.disableButton(this.buttonSyncQAs, true);

    try {
      await Excel.run(async context => {
        context.workbook.worksheets.load("items");

        await context.sync();

        let allRanges = [];
        context.workbook.worksheets.items.forEach((element, index) => {
          if (index == 0) {
            return;
          }
          // TODO max 4 lines
          let range = element.getRange("A1:B4");
          range.load("values");
          allRanges.push(range);
        });

        await context.sync();

        let config = await this.getConfig(context);

        let qaList = [];
        allRanges.forEach(element => {
          let enabled: boolean = false;
          let endpoint = new QnAMakerEndpoint();
          endpoint.EndpointKey = config.get(ConfigKeys.EndpointKey);
          endpoint.Host = config.get(ConfigKeys.Host);

          for (let i = 0;i < element.values.length;++i) {
            let ele = element.values[i];
            // TODO break when blank line
            if (ele.length < 2 || String(ele[0]) == "") {
              break;
            }
            switch (String(ele[0]).toLowerCase()) {
              case ConfigKeys.Id:
                endpoint.KnowledgeBaseId = String(ele[1]);
                break;
              case ConfigKeys.Status:
                if (String(ele[1]).toLowerCase() == Status.Enable) {
                  enabled = true;
                }
                break;
            }
          }

          if (enabled) {
            qaList.push(endpoint);
          }
        });
        if (qaList.length > 0) {
          this.pushEvent(Event.SetQnA, qaList);
        } else {
          this.disableButton(this.buttonSyncQAs, false);
        }
      });

      // TODO why
      setTimeout(() => {this.clickDoSync()}, 1000);
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickSyncQA = async () => {
    this.disableButton(this.buttonSyncQA, true);
    try {
      await Excel.run(async context => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('position');
        sheet.load('name');

        let range = sheet.getUsedRange();
        range.load('values');

        await context.sync();

        if (sheet.position == 0) return;

        let data = new Map<string, QnADTO>();
        let dataStart = false;
        let id = null;
        let lastKey: string = null;
        range.values.forEach(element => {
          if (element.length < 2) return;
          // value is question, key is answer
          let value = String(element[0]);
          let key = String(element[1]);
          switch (value.toLowerCase()) {
            case ConfigKeys.Id:
              id = key;
              break;
            case "":
              // TODO use blank to separate config and data
              dataStart = true;
              return;
          }
          if (!dataStart) return;
          
          // use last answer if empty
          if (key == "") {
            key = lastKey;
          }
          if (data.has(key)) {
            data.get(key).questions.push(value);
          } else {
            data.set(key, new QnADTO(key, Source.Editorial, value));
          }
          lastKey = key;
        });
        if (data.size == 0) {
          return;
        }
        this.addDebug(`Total QA: ${data.size}`);

        let key = String((await this.getConfig(context)).get(ConfigKeys.SubscriptionKey));
        const url = `https://westus.api.cognitive.microsoft.com/qnamaker/v4.0/knowledgebases/${id}`;
        let response = await axios.put(url,
          {
            "qnAList": Array.from(data.values())
          },
          { 
            headers: {
              'Content-Type': 'application/json',
              'Ocp-Apim-Subscription-Key': key
            }
          });
        this.addDebug(`Replacing data status: ${response.status}. Wait for publishing..`);
        response = await axios.post(url,
          {},
          { 
            headers: {
              'Ocp-Apim-Subscription-Key': key
            }
          });
        this.addDebug(`Publishing status: ${response.status}. Wait for updating name..`);
        response = await axios.patch(url,
          {
            "update": {
              "name": sheet.name
            }
          },
          {
            headers: {
              'Content-Type': 'application/json',
              'Ocp-Apim-Subscription-Key': key
            }
          });
        this.addDebug(`Updating status: ${response.status}.`);
      });

      this.disableButton(this.buttonSyncQA, false);
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickCreateQA = async () => {
    this.disableButton(this.buttonCreateQA, true);
    try {
      await Excel.run(async context => {
        let config = await this.getConfig(context);

        const url = `https://westus.api.cognitive.microsoft.com/qnamaker/v4.0/knowledgebases/create`;
        let response = await axios.post(url,
          {
            "name": this.newQAName
          },
          {
            headers: {
              'Content-Type': 'application/json',
              'Ocp-Apim-Subscription-Key': String(config.get(ConfigKeys.SubscriptionKey))
            }
          });
        this.addDebug(`Creating status: ${response.status}. Wait for finishing..`);
        this.checkCreateQA(config, response.data.operationId);
      });
    } catch(error) {
      this.addDebug(error);
    }
  };

  async checkCreateQA(config: Map<any, any>, operationId: string) {
    const url = `https://westus.api.cognitive.microsoft.com/qnamaker/v4.0/operations/${operationId}`;
    let response = await axios.get(url,
      {
        headers: {
          'Content-Type': 'application/json',
          'Ocp-Apim-Subscription-Key': String(config.get(ConfigKeys.SubscriptionKey))
        }
      });
    if (response.data.operationState != 'Succeeded') {
      setTimeout(() => { this.checkCreateQA(config, operationId); }, 1000);
      return;
    }

    try {
      await Excel.run(async context => {
        let id = String(response.data.resourceLocation).split('/')[2];

        let sheet = context.workbook.worksheets.add(this.newQAName);
        let range = sheet.getRange("A1:B2");
        range.values = [[ConfigKeys.Status, "disable"],[ConfigKeys.Id, id]];
        sheet.activate();

        await context.sync();

        this.addDebug(`Finishing status ${response.status}. Add QAs and click 'Sync QA' for the new QA.`);
      });

      this.disableButton(this.buttonCreateQA, false);
    } catch (error) {
      this.addDebug(error);
    }
  }

  clickDoSync = async () => {
    // const toDispatch = this.toDispatch.length;
    this.toDispatch.forEach(element => {
      this.props.store.dispatch(element);
    });
    this.toDispatch = [];
    this.disableButton(this.buttonSyncConfig, false);
    this.disableButton(this.buttonSyncQAs, false);
    // TODO why?
    // setTimeout(() => {this.addDebug(`Sent ${toDispatch} configs.`);}, 1000);
  };

  render() {
    const { title, isOfficeInitialized, store } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    this.getTokenId();

    return (
      <div className="ms-welcome">
        <div>Debug<button onClick={this.clickClearDebug}>Clear Debug</button></div>
        <div>{this.state.debugstring.map((value, index) => {
          return (<p>{index}: {value}</p>)
        })}</div>
        <div>
          <button id={this.buttonSyncConfig} onClick={this.clickSyncConfig}>Sync Config</button>
          <button id={this.buttonSyncQAs} onClick={this.clickSyncQAs}>Sync QAs</button>
          <button id={this.buttonSyncQA} onClick={this.clickSyncQA}>Sync QA</button>
          <button id={this.buttonCreateQA} onClick={this.clickCreateQA}>Create QA</button>
          <button id='DoSync' onClick={this.clickDoSync}>Do Sync</button>
        </div>
        {this.state.webChatToken &&
          <ReactWebChat directLine={new DirectLine({ token: this.state.webChatToken })} userID={this.tempUserId} store={store}/>
        }
      </div>
    );
  }
}
