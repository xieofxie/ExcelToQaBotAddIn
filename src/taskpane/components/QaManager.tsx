import * as React from 'react';
import { Collapse } from 'react-collapse';
import { AddSourceEvent, DelSourceEvent, Event, QnAMakerEndpoint, QnAMakerEndpointEx, Source } from '../models/Event';
import { QnAMakerModels } from '@azure/cognitiveservices-qnamaker';
import { Element } from 'react-scroll'
import axios from "axios";

export interface QaManagerProps {
    qnAs: { [index: string]: QnAMakerEndpointEx },
    fileHostUrl?: string,
    syncToThis: Function,
    pushEvent: Function,
    clickDoSync: Function,
    addDebug: Function,
}

interface QaManagerState {
    isOpened: boolean;
    // TODO tired of finding working resizable
    height: number;
}

function InputWithButtonComponent(props: { init:string, button:string, onClick:(name:string)=>void }){
    const [ name, setName ] = React.useState(props.init);
    return (<span>
        <input type="text" value={name} onChange={e=>{setName(e.target.value)}} />
        <button onClick={()=>{props.onClick(name)}}>{props.button}</button>
    </span>)
}

export class QaManager extends React.Component<QaManagerProps, QaManagerState> {
    buttonCreateQA = "CreateQA";
    buttonGetQAs = "GetQAs";
    newQAName = "New QA";

    constructor(props: QaManagerProps, context: any) {
        super(props, context);
        this.state = {
            isOpened: true,
            height: 200,
        };
    }

    componentDidMount() {
        // TODO why?
        // this.clickGetQAs();
    }

    SourceItem = (props: {knowledgeBaseId: string, source: Source}) => {
        return (<div>
            Id: {props.source.id} Description: {props.source.description}
            <button onClick={() => {this.clickDeleteSource(props.knowledgeBaseId, props.source)}}>Delete</button>
        </div>)
    };

    QnAItem = (props: {qnA: QnAMakerEndpointEx}) => {
        const { fileHostUrl } = this.props;
        const { qnA } = props;
        const divStyle = {
            border: '1px solid black'
        };
        return (<div style={divStyle}>
            <div>
                Name: <InputWithButtonComponent init={qnA.name} button='Update' onClick={(name)=>{this.clickUpdateName(qnA.knowledgeBaseId, name)}}/>
                {qnA.enable?'Enabled':'Disabled'}
                <button onClick={()=>{this.clickToggleEnable(qnA.knowledgeBaseId, !qnA.enable)}}>{qnA.enable?'Disable':'Enable'}</button>
                <button onClick={()=>{this.clickSyncToThis(qnA.knowledgeBaseId)}}>Add/Update To This</button>
                {fileHostUrl && <label>Add/Update File To This<input type="file" onChange={(e)=>{this.clickAddFile(qnA.knowledgeBaseId, e)}}/></label>}
                <button onClick={()=>{this.clickDeleteQA(qnA.knowledgeBaseId)}}>Delete</button>
            </div>
            {Object.values(qnA.sources).map(v => {
                return <this.SourceItem key={v.id} knowledgeBaseId={qnA.knowledgeBaseId} source={v}/>;
            })}
        </div>);
    };

    clickAddFile = async (knowledgeBaseId: string, event: React.ChangeEvent<HTMLInputElement>) => {
        const { fileHostUrl } = this.props;
        if (!fileHostUrl) return;
        if (event.target.files == null) return;
        const file = event.target.files[0];
        let response = null;
        try {
            response = await axios.post(fileHostUrl,
                file,
                {
                    headers: {
                        'Content-Type': file.type
                    },
                    params: {
                        name: file.name,
                        type: file.type
                    }
                });
        } catch (error) {
            //this.props.addDebug(error);
        }
        if (response == null) return;

        let value: AddSourceEvent = {
            knowledgeBaseId: knowledgeBaseId,
            filesDescription: [`${file.name} ${String(file.size)}`],
            files: [{
                fileName: file.name,
                fileUri: `${fileHostUrl}?id=${String(response.data)}`
            }]
        };
        this.props.pushEvent(Event.AddSource, value, true);
    };

    clickCreateQA = (name: string) => {
        const value: QnAMakerModels.CreateKbDTO = {
            name: name
        };
        value.name = name;
        this.props.pushEvent(Event.CreateQnA, value, true);
    }

    clickGetQAs = () => {
        this.props.pushEvent(Event.GetQnA, null, true);
    };

    clickUpdateName = async (knowledgeBaseId: string, name: string) => {
        const value = new QnAMakerEndpointEx();
        value.knowledgeBaseId = knowledgeBaseId;
        value.name = name;
        this.props.pushEvent(Event.UpdateQnA, value, true);
    };

    clickToggleEnable = async (knowledgeBaseId: string, enable: boolean) => {
        const value = new QnAMakerEndpointEx();
        value.knowledgeBaseId = knowledgeBaseId;
        value.enable = enable;
        this.props.pushEvent(Event.EnableQnA, value, true);
    }

    clickSyncToThis = async (knowledgeBaseId: string) => {
        this.props.syncToThis(knowledgeBaseId);
    };

    clickDeleteQA = (knowledgeBaseId: string) => {
        const value = new QnAMakerEndpoint();
        value.knowledgeBaseId = knowledgeBaseId;
        this.props.pushEvent(Event.DelQnA, value, true);
    };

    clickDeleteSource = async (knowledgeBaseId: string, source: Source) => {
        const value = new DelSourceEvent();
        value.knowledgeBaseId = knowledgeBaseId;
        value.ids = [source.id];
        this.props.pushEvent(Event.DelSource, value, true);
    }

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
                        <InputWithButtonComponent init={this.newQAName} button='Create QA' onClick={name=>{this.clickCreateQA(name)}} />
                        <button id={this.buttonGetQAs} onClick={this.clickGetQAs}>Get QAs</button>
                    </div>
                    <div style={{height}}>
                        <Element name='QaManagerElement' style={{position: 'relative', height: '100%', overflow: 'scroll'}}>
                            {qnAs && Object.values(qnAs).map(v => { return <this.QnAItem key={v.knowledgeBaseId} qnA={v} /> })}
                        </Element>
                    </div>
                </Collapse>
            </div>
        )
    }
}
