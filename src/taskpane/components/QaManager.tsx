import * as React from 'react';
import { Collapse } from 'react-collapse';
import { EnableQnAEvent, Event, QnADTO, QnAMakerEndpointEx, Source, SourceEvent, SourceType } from '../models/Event';
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

    SourceItem = (props) => {
        return (<div>
            Id: {props.source.Id} Description: {props.source.Description}
            <button onClick={() => {this.clickDeleteSource(props.knowledgeBaseId, props.source)}}>Delete</button>
        </div>)
    };

    renderQnA = (qnA: QnAMakerEndpointEx) => {
        const divStyle = {
            border: '1px solid black'
        };
        return (<div style={divStyle}>Name: {qnA.name} Enabled: {String(qnA.enable)}
            <div>
                <button onClick={()=>{this.clickToggleEnable(qnA.knowledgeBaseId, !qnA.enable)}}>{qnA.enable?'Disable':'Enable'}</button>
                <button onClick={()=>{this.clickSyncToThis(qnA.knowledgeBaseId)}}>Sync To This</button>
            </div>
            {Object.values(qnA.sources).map(v => {
                return <this.SourceItem key={v.Id} knowledgeBaseId={qnA.knowledgeBaseId} source={v}/>;
            })}
        </div>);
    };

    clickGetQAs = () => {
        try {
            this.props.pushEvent(Event.GetQnA, null);
      
            // TODO why
            setTimeout(() => {this.props.clickDoSync()}, 1000);
        } catch (error) {
            this.props.addDebug(error);
        }
    };

    clickToggleEnable = async (knowledgeBaseId: string, enable: boolean) => {
        const value = new EnableQnAEvent(knowledgeBaseId, enable);
        this.props.pushEvent(Event.EnableQnA, value);
        setTimeout(() => this.props.clickDoSync(), 1000);
    }

    clickSyncToThis = async (knowledgeBaseId: string) => {
        try {
            await Excel.run(async context => {
                let book = context.workbook;
                book.load('name');

                let sheet = book.worksheets.getActiveWorksheet();
                sheet.load('position');
                sheet.load('name');

                let range = sheet.getUsedRange();
                range.load('values');

                await context.sync();

                if (sheet.position == 0) return;

                let data = new Map<string, QnADTO>();
                let lastKey: string = null;
                range.values.forEach(element => {
                    if (element.length < 2) return;
                    // value is question, key is answer
                    let value = String(element[0]);
                    let key = String(element[1]);
                    // use last answer if empty
                    if (key == "") {
                        key = lastKey;
                    }
                    if (data.has(key)) {
                        data.get(key).questions.push(value);
                    } else {
                        data.set(key, new QnADTO(key, value));
                    }
                    lastKey = key;
                });
                if (data.size == 0) {
                    return;
                }
                this.props.addDebug(`Total QA: ${data.size}`);

                let value = new SourceEvent();
                value.KnowledgeBaseId = knowledgeBaseId;
                value.QnaList = Array.from(data.values());
                value.Id = sheet.name;
                value.Description = book.name;
                value.Type = SourceType.Editorial;

                this.props.pushEvent(Event.AddSource, value);
            });

            setTimeout(() => {this.props.clickDoSync()}, 1000);
        } catch (error) {
            this.props.addDebug(error);
        }
    };

    clickDeleteSource = async (knowledgeBaseId: string, source: Source) => {
        const value = new SourceEvent();
        value.KnowledgeBaseId = knowledgeBaseId;
        value.Id = source.Id;
        value.Type = source.Type;
        this.props.pushEvent(Event.DelSource, value);
        setTimeout(() => this.props.clickDoSync(), 1000);
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
