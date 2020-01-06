import * as React from 'react';
import ReactWebChat from 'botframework-webchat';
import { DirectLine } from 'botframework-directlinejs';
import {Collapse} from 'react-collapse';
// TODO only found this working
import { Element } from 'react-scroll'

export interface WebChatProps {
    webChatToken: string;
    tempUserId: string;
    store: any;
}

interface WebChatState {
    isOpened: boolean;
    // TODO tired of finding working resizable
    height: number;
}

export class WebChat extends React.Component<WebChatProps, WebChatState> {
    constructor(props: WebChatProps, context: any) {
        super(props, context);
        this.state = {
            isOpened: true,
            height: 100
        };
    }

    render() {
        const {isOpened, height} = this.state;
        const {webChatToken, tempUserId, store} = this.props;

        return (
            <div>
                <div>
                    <div style={{width: "100px", float:"left"}}>
                        <label>WebChat</label>
                        <input
                            type="checkbox"
                            checked={isOpened}
                            onChange={({target: {checked}}) => this.setState({isOpened: checked})} />
                    </div>
                    <div style={{marginLeft: "100px", marginRight: "20px"}}>
                        <input
                            type="range"
                            value={height}
                            step={20}
                            min={100}
                            max={1000}
                            style={{width:"100%"}}
                            onChange={({target: {value}}) => this.setState({height: parseInt(value, 10)})} />
                    </div>
                </div>
                <Collapse isOpened={isOpened}>
                    <div style={{height}}>
                        <Element style={{position: 'relative', height: '100%', overflow: 'scroll'}}>
                            <ReactWebChat className='WebChat' directLine={new DirectLine({ token: webChatToken })} userID={tempUserId} store={store}/>
                        </Element>
                    </div>
                </Collapse>
            </div>
        );
    }
}
