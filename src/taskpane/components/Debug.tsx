import * as React from 'react';
import {Collapse} from 'react-collapse';
// TODO only found this working
import { Element } from 'react-scroll'
import {v1 as uuid} from 'uuid';

export interface AppProps {
    debugString: string[];
    clearCb: Function;
}

interface AppState {
    isOpened: boolean;
    // TODO tired of finding working resizable
    height: number;
}

export class Debug extends React.Component<AppProps, AppState> {
    uuid: string;

    constructor(props: AppProps, context: any) {
        super(props, context);
        this.state = {
            isOpened: true,
            height: 100
        };
        this.uuid = uuid();
    }

    clickClearDebug = async() => {
        this.props.clearCb();
    };

    componentDidMount() {
    }

    renderItem(index) {
        return (<p>{index}: {this.props.debugString[index]}</p>)
    }

    render() {
        const {isOpened, height} = this.state;

        return (
            <div>
                <div>
                    Debug
                    <input type="checkbox" checked={isOpened} onChange={({target: {checked}}) => this.setState({isOpened: checked})} />
                    <input
                        type="range"
                        value={height}
                        step={50}
                        min={100}
                        max={1000}
                        onChange={({target: {value}}) => this.setState({height: parseInt(value, 10)})} />
                </div>
                <Collapse isOpened={isOpened}>
                    <div>
                        <button onClick={this.clickClearDebug}>Clear Debug</button>
                    </div>
                    <div style={{height}}>
                        <Element style={{position: 'relative', height: '100%', overflow: 'scroll'}}>
                            {this.props.debugString.map((value, index) => {
                                return (<p>{index}: {value}</p>);
                            })}
                        </Element>
                    </div>
                </Collapse>
            </div>
        );
    }
}