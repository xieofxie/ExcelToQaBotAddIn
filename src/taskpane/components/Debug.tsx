import * as React from 'react';
import {Collapse} from 'react-collapse';
// TODO only found this working
import { Element } from 'react-scroll'
import {v1 as uuid} from 'uuid';

export interface DebugProps {
    debugString: string[];
    clearCb: Function;
}

interface DebugState {
    isOpened: boolean;
    // TODO tired of finding working resizable
    height: number;
}

export class Debug extends React.Component<DebugProps, DebugState> {
    uuid: string;

    constructor(props: DebugProps, context: any) {
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

    renderItem = (index) => {
        return (<p>{index}: {this.props.debugString[index]}</p>)
    };

    render() {
        const {isOpened, height} = this.state;

        return (
            <div>
                <div>
                    <div style={{width: "100px", float:"left"}}>
                        <label>Debug</label>
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
                    <div>
                        <button onClick={this.clickClearDebug}>Clear Debug</button>
                    </div>
                    <div style={{height}}>
                        <Element style={{position: 'relative', height: '100%', overflow: 'scroll'}}>
                            {this.props.debugString.map((value, index) => {
                                return (<p key={index}>{index}: {value}</p>);
                            })}
                        </Element>
                    </div>
                </Collapse>
            </div>
        );
    }
}