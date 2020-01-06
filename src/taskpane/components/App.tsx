import * as React from "react";
import Progress from "./Progress";
import { DirectLine } from 'botframework-directlinejs';
import ReactWebChat from 'botframework-webchat';
import { ConfigKeys } from '../models/Config';
import { Event } from '../models/Event';
import { Debug } from "./Debug";
import { LgEditor } from "./LgEditor";
import { QaManager } from "./QaManager";
import { getConfig } from "../utils/Utils";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
  store: any;
  eventDispatcher: any;
}

export interface AppState {
  webChatToken: string;
  debugstring: string[];
  qnAs: {};
}

export default class App extends React.Component<AppProps, AppState> {
  tempUserId = 'TempUserId';
  toDispatch = [];
  eventListener: string;
  buttonSyncConfig = 'SyncConfig';

  constructor(props: AppProps, context: any) {
    super(props, context);
    this.state = {
      webChatToken: '',
      debugstring: [],
      qnAs: {},
    };
  }

  componentDidMount() {
    this.eventListener = this.props.eventDispatcher.register((event) => {
      if (event.name == Event.GetQnA) {
        this.setState({ qnAs: event.value });
      }
    });
  }

  componentWillUnmount() {
    this.props.eventDispatcher.unregister(this.eventListener);
  }

  addDebug = (error: any) => {
    this.setState((state) => {
      return { debugstring: state.debugstring.concat(String(error)) }
    });
  };

  clearDebug = () => {
    this.setState({debugstring: []});
  };

  pushEvent = (name: string, value: any) => {
    this.toDispatch.push({
      type: 'WEB_CHAT/SEND_EVENT',
      payload: {name: name, value: value}
    });

    // this.addDebug(`${name}:${value}`);
  };

  getTokenId = async () => {
    try {
      // TODO why?
      if (this.state.webChatToken != '') {
        return;
      }

      await Excel.run(async context => {
        let token = String((await getConfig(context)).get(ConfigKeys.WebChatToken));
        this.setState({ webChatToken: token });
        this.addDebug(token);
      });
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickSyncConfig = async() => {
    try {
      await Excel.run(async context => {
        let config = await getConfig(context);

        config.forEach((value, key) => {
          switch (key) {
            case ConfigKeys.ResultNumber:
              this.pushEvent(Event.SetResultNumber, Number(value));
              break;
            case ConfigKeys.MinScore:
              this.pushEvent(Event.SetMinScore, Number(value));
              break;
            case ConfigKeys.Debug:
              this.pushEvent(Event.SetDebug, Boolean(value));
              break;
          }
        });
      });

      // TODO why
      setTimeout(() => {this.clickDoSync()}, 1000);
    } catch (error) {
      this.addDebug(error);
    }
  };

  clickDoSync = async () => {
    // const toDispatch = this.toDispatch.length;
    this.toDispatch.forEach(element => {
      this.props.store.dispatch(element);
    });
    this.toDispatch = [];

    // TODO why? maybe it causes a refresh of UI which breaks directline
    // setTimeout(() => {this.addDebug(`Sent ${toDispatch} configs.`);}, 1000);
  };

  render() {
    const { title, isOfficeInitialized, store } = this.props;
    const { webChatToken, debugstring, qnAs } = this.state;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    this.getTokenId();

    if (webChatToken == '') {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Wait for reading token." />
      );
    }

    return (
      <div>
        <Debug debugString={debugstring} clearCb={this.clearDebug}/>
        <LgEditor
          pushEvent={this.pushEvent}
          clickDoSync={this.clickDoSync}
          addDebug={this.addDebug}
        />
        <QaManager
          qnAs={qnAs}
          pushEvent={this.pushEvent}
          clickDoSync={this.clickDoSync}
          addDebug={this.addDebug}
        />
        <div>
          <button id={this.buttonSyncConfig} onClick={this.clickSyncConfig}>Sync Config</button>
          <button id='DoSync' onClick={this.clickDoSync}>Do Sync</button>
        </div>
        <ReactWebChat directLine={new DirectLine({ token: webChatToken })} userID={this.tempUserId} store={store}/>
      </div>
    );
  }
}
