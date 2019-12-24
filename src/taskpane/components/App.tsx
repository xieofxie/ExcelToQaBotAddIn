import * as React from "react";
import axios from "axios";
import Progress from "./Progress";
import { DirectLine } from 'botframework-directlinejs';
import ReactWebChat, { createStore } from 'botframework-webchat';
import { Config, Status } from '../models/Config';
import { Event, QnAMakerEndpoint } from '../models/Event';
import { QnADTO, Source } from "../models/QnAMaker";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  token: string;
  debugstring: string[];
}

export default class App extends React.Component<AppProps, AppState> {

  tempUserId = 'TempUserId';
  toDispatch = [];
  buttonSyncConfig = 'SyncConfig';
  buttonSyncQAs = 'SyncQAs';
  buttonSyncQA = 'SyncQA';

  store = createStore();

  constructor(props: AppProps, context: any) {
    super(props, context);
    this.state = {
      token: '',
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

    this.addDebug(`${name}:${value}`);
  }

  disableButton(name: string, disable: boolean) {
    (document.getElementById(name) as HTMLButtonElement).disabled = disable;
  }

  async getOneConfig(context: Excel.RequestContext, config: string) {
    const configSheet = context.workbook.worksheets.getFirst();
    const configRange = configSheet.getUsedRange();
    configRange.load("values");

    await context.sync();

    for (let i = 0;i < configRange.values.length;++i) {
      let element = configRange.values[i];
      if (element.length < 2) continue;
      if (String(element[0]).toLowerCase() == config) {
        return element[1];
      }
    }
  }

  getTokenId = async () => {
    try {
      // TODO why?
      if (this.state.token != '') {
        return;
      }

      await Excel.run(async context => {
        let token = String(await this.getOneConfig(context, Config.Token));
        this.setState({ token: token });
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
        const configSheet = context.workbook.worksheets.getFirst();
        const configRange = configSheet.getUsedRange();
        configRange.load("values");

        await context.sync();

        let pushed = false;
        configRange.values.forEach(element => {
          if (element.length < 2) return;
          switch (String(element[0]).toLowerCase()) {
            case Config.ResultNumber:
              this.pushEvent(Event.SetResultNumber, Number(element[1]));
              pushed = true;
              break;
            case Config.NoResultResponse:
              this.pushEvent(Event.SetNoResultResponse, String(element[1]));
              pushed = true;
              break;
            case Config.MinScore:
              this.pushEvent(Event.SetMinScore, Number(element[1]));
              pushed = true;
              break;
            case Config.Debug:
              this.pushEvent(Event.SetDebug, Boolean(element[1]));
              pushed = true;
              break;
          }
        });
        if (!pushed) {
          this.disableButton(this.buttonSyncConfig, false);
        }
      });
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

        let qaList = [];
        allRanges.forEach(element => {
          let enabled: boolean = false;
          let endpoint = new QnAMakerEndpoint();

          for (let i = 0;i < element.values.length;++i) {
            let ele = element.values[i];
            // TODO break when blank line
            if (ele.length < 2 || String(ele[0]) == "") {
              break;
            }
            switch (String(ele[0]).toLowerCase()) {
              case Config.Id:
                endpoint.KnowledgeBaseId = String(ele[1]);
                break;
              case Config.Key:
                endpoint.EndpointKey = String(ele[1]);
                break;
              case Config.Host:
                endpoint.Host = String(ele[1]);
                break;
              case Config.Status:
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

        await context.sync();

        if (sheet.position == 0) return;
        let range = sheet.getUsedRange();
        range.load('values');

        await context.sync();

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
            case Config.Id:
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
          this.disableButton(this.buttonSyncQA, false);
          return;
        }
        this.addDebug(data.size);

        let key = String(await this.getOneConfig(context, Config.Key));
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
        this.addDebug(response.status);
        response = await axios.post(url,
          {},
          { 
            headers: {
              'Ocp-Apim-Subscription-Key': key
            }
          });
        this.addDebug(response.status);
        this.disableButton(this.buttonSyncQA, false);
      });
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickDoSync = async () => {
    this.toDispatch.forEach(element => {
      this.store.dispatch(element);
    });
    this.toDispatch = [];
    this.disableButton(this.buttonSyncConfig, false);
    this.disableButton(this.buttonSyncQAs, false);
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

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
          <button id='DoSync' onClick={this.clickDoSync}>Do Sync</button>
        </div>
        {this.state.token &&
          <ReactWebChat directLine={new DirectLine({ token: this.state.token })} userID={this.tempUserId} store={this.store}/>
        }
      </div>
    );
  }
}
