import * as React from "react";
import Progress from "./Progress";
import { DirectLine } from 'botframework-directlinejs';
import ReactWebChat, { createStore } from 'botframework-webchat';
import Config from '../models/Config';
import Event from '../models/Event';
import QnAMakerEndpoint from '../models/QnAMakerEndpoint';
import Status from '../models/Status';
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

  store = createStore();

  constructor(props, context) {
    super(props, context);
    this.state = {
      token: '',
      debugstring: []
    };
  }

  addDebug(error) {
    this.setState({debugstring: this.state.debugstring.concat(String(error))});
  }

  clearDebug() {
    this.setState({debugstring: []});
  }

  pushEvent(name, value) {
    this.toDispatch.push({
      type: 'WEB_CHAT/SEND_EVENT',
      payload: {name: name, value: value}
    });

    this.addDebug(`${name}:${value}`);
  }

  clickClearDebug = async () => {
    this.clearDebug();
  }

  componentDidMount() {
  }

  getTokenId = async () => {
    try {
      // TODO why?
      if (this.state.token != '') {
        return;
      }

      await Excel.run(async context => {
        const configSheet = context.workbook.worksheets.getFirst();
        const configRange = configSheet.getUsedRange();
        configRange.load("values");

        await context.sync();

        for (let i = 0;i < configRange.values.length;++i) {
          let element = configRange.values[i];
          if (element.length < 2) continue;
          if (String(element[0]).toLowerCase() == Config.Token) {
            this.setState({ token: String(element[1]) });
            this.addDebug(String(element[1]));
            break;
          }
        }
      });
    } catch (error) {
      this.addDebug(error);
    }
  }

  clickSyncConfig = async() => {
    try {
      (document.getElementById('SyncConfig') as HTMLButtonElement).disabled = true;

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
          (document.getElementById('SyncConfig') as HTMLButtonElement).disabled = false;
        }
      });
    } catch (error) {
      this.addDebug(error);
    }
  }

  clickSyncQA = async () => {
    (document.getElementById('SyncQA') as HTMLButtonElement).disabled = true;

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
              case Config.key:
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
          (document.getElementById('SyncQA') as HTMLButtonElement).disabled = false;
        }
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
    (document.getElementById('SyncConfig') as HTMLButtonElement).disabled = false;
    (document.getElementById('SyncQA') as HTMLButtonElement).disabled = false;
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
          <button id='SyncConfig' onClick={this.clickSyncConfig}>Sync Config</button>
          <button id='SyncQA' onClick={this.clickSyncQA}>Sync QA</button>
          <button id='DoSync' onClick={this.clickDoSync}>Do Sync</button>
        </div>
        {this.state.token &&
          <ReactWebChat directLine={new DirectLine({ token: this.state.token })} userID={this.tempUserId} store={this.store}/>
        }
      </div>
    );
  }
}
