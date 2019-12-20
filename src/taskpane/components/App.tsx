import * as React from "react";
import Progress from "./Progress";
import { DirectLine } from 'botframework-directlinejs';
import ReactWebChat, { createStore } from 'botframework-webchat';
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  token: string;
  debugstring: string;
}

class QnAMakerEndpoint {
  KnowledgeBaseId;
  EndpointKey;
  Host;

  constructor(id, key, host) {
    this.KnowledgeBaseId = id;
    this.EndpointKey = key;
    this.Host = host;
  }
}

const SetQnA = "SetQnA";

const SetResultNumber = "SetResultNumber";

//const SetMinScore = "SetMinScore";

const SetNoResultResponse = "SetNoResultResponse";

export default class App extends React.Component<AppProps, AppState> {

  tempUserId = 'TempUserId';
  toDispatch = [];

  store = createStore();

  constructor(props, context) {
    super(props, context);
    this.state = {
      token: '',
      debugstring: ''
    };
  }

  addDebug(error) {
    const debugstring = this.state.debugstring == '' ? `${error}` : this.state.debugstring + `\n${error}`;
    this.setState({debugstring: debugstring});
  }

  clearDebug() {
    this.setState({debugstring: ""});
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
        const configRange = configSheet.getRange("B1");
        configRange.load("values");

        await context.sync();

        this.setState({ token: String(configRange.values[0][0]) });
        this.addDebug(configRange.values[0][0]);
      });
    } catch (error) {
      this.addDebug(error);
    }
  }

  clickSyncConfig = async() => {
    try {
      await Excel.run(async context => {
        const configSheet = context.workbook.worksheets.getFirst();
        const configRange = configSheet.getRange("B2:B3");
        configRange.load("values");

        await context.sync();

        this.toDispatch.push({
          type: 'WEB_CHAT/SEND_EVENT',
          payload: { name: SetResultNumber, value: Number(configRange.values[0][0]) }
        });
        this.toDispatch.push({
          type: 'WEB_CHAT/SEND_EVENT',
          payload: { name: SetNoResultResponse, value: String(configRange.values[1][0]) }
        });
        this.addDebug(`${configRange.values[0][0]},${configRange.values[1][0]}`);
      });
    } catch (error) {
      this.addDebug(error);
    }
  }

  clickSyncQA = async () => {
    try {
      await Excel.run(async context => {
        context.workbook.worksheets.load("items");

        await context.sync();

        let allRanges = [];
        context.workbook.worksheets.items.forEach((element, index) => {
          if (index == 0) {
            return;
          }
          let range = element.getRange("B1:B4");
          range.load("values");
          allRanges.push(range);
        });

        await context.sync();

        let qaList = [];
        allRanges.forEach(element => {
          if (String(element.values[0][0]).toLowerCase() != "enable") {
            return;
          }
          qaList.push(new QnAMakerEndpoint(String(element.values[1][0]), String(element.values[2][0]), String(element.values[3][0])));
        });
        if (qaList.length > 0) {
          this.toDispatch.push({
            type: 'WEB_CHAT/SEND_EVENT',
            payload: { name: SetQnA, value: qaList }
          })
        }
        this.addDebug(qaList.length);
      });
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickDoSync = async () => {
    this.toDispatch.forEach(element => {
      this.store.dispatch(element);
    });
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
        <div><p>{this.state.debugstring}</p></div>
        <div>
          <button onClick={this.clickSyncConfig}>Sync Config</button>
          <button onClick={this.clickSyncQA}>Sync QA</button>
          <button onClick={this.clickDoSync}>Do Sync</button>
        </div>
        {this.state.token &&
          <ReactWebChat directLine={new DirectLine({ token: this.state.token })} userID={this.tempUserId} store={this.store}/>
        }
      </div>
    );
  }
}
