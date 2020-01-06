import * as React from 'react';
import {Collapse} from 'react-collapse';
import {EnableQnAEvent, Event, QnAMakerEndpointEx} from '../models/Event';
import { Element } from 'react-scroll'

export interface QaManagerProps {
    qnAs: { [index: string]: QnAMakerEndpointEx },
    pushEvent: Function,
    clickDoSync: Function,
    addDebug: Function,
}

interface QaManagerState {
    isOpened: boolean;
    // TODO tired of finding working resizable
    height: number;
}

export class QaManager extends React.Component<QaManagerProps, QaManagerState> {
    buttonCreateQA = "CreateQA";
    buttonGetQAs = "GetQAs";
    newQAName = "New QA";

    constructor(props: QaManagerProps, context: any) {
        super(props, context);
        this.state = {
            isOpened: true,
            height: 100,
        };
    }

    componentDidMount() {
        // TODO why?
        // this.clickGetQAs();
    }

    renderQnA(qnA: QnAMakerEndpointEx) {
        const divStyle = {
            border: '1px solid black'
        };
        return (<div style={divStyle}>Name: {qnA.name} Enabled: {String(qnA.enable)}
            <div>
                <button onClick={()=>{
                const value = new EnableQnAEvent(qnA.knowledgeBaseId, !qnA.enable);
                this.props.pushEvent(Event.EnableQnA, value);
                setTimeout(() => this.props.clickDoSync(), 1000);
                }}>{qnA.enable?'Disable':'Enable'}</button>
            </div>
            {Object.values(qnA.sources).map(v => {
            return (<div key={v.Id}>Id: {v.Id} Description: {v.Description}</div>);
            })}
        </div>);
    }

    clickGetQAs = () => {
        try {
            this.props.pushEvent(Event.GetQnA, null);
      
            // TODO why
            setTimeout(() => {this.props.clickDoSync()}, 1000);
        } catch (error) {
            this.props.addDebug(error);
        }
    };

    render() {
        const { qnAs } = this.props;
        const { isOpened, height } = this.state;

        return (
            <div>
                <div>
                    <div style={{width: "100px", float:"left"}}>
                        <label>Qa Manager</label>
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
                        <button id={this.buttonCreateQA} onClick={()=>{}}>Create QA</button>
                        <button id={this.buttonGetQAs} onClick={this.clickGetQAs}>Get QAs</button>
                    </div>
                    <div style={{height}}>
                        <Element style={{position: 'relative', height: '100%', overflow: 'scroll'}}>
                            {Object.values(qnAs).map(v => { return this.renderQnA(v) })}
                        </Element>
                    </div>
                </Collapse>
            </div>
        )
    }
}
