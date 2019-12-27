import * as React from "react";
import MonacoEditor from 'react-monaco-editor';
import {Collapse} from 'react-collapse';
import * as monacoEditor from 'monaco-editor/esm/vs/editor/editor.api';
import { registerLGLanguage } from './lg';
import { throttle } from 'lodash'
import { Event } from '../models/Event';
import { ConfigKeys } from "../models/Config";
import { getConfig } from "../utils/Utils";

export interface LgEditorProps {
    pushEvent: Function,
    clickDoSync: Function,
    disableButton: Function,
    addDebug: Function
}

interface LgEditorState {
    isOpened: boolean;
    height: number;
    width: number;
    code: string;
}

export class LgEditor extends React.Component<LgEditorProps, LgEditorState> {

    myRef = null;
    buttonTest = "LgEditorTest";
    buttonSaveToExcel = "LgEditorSaveToExcel";
    buttonSyncLg = "LgEditorSyncLg";

    constructor(props: LgEditorProps, context: any) {
        super(props, context);
        this.state = {
            isOpened: !false,
            height: 500,
            width: 500,
            code: "# CreateAnswer(results, debug)\r\n- Sample answer."
        };
        this.myRef = React.createRef();
    }

    updateWidth = throttle(() => {
        const width = this.myRef.current.getBoundingClientRect().width;
        if (width != this.state.width) {
            this.setState({width: width});
        }
    }, 100);

    componentWillMount = async () => {
        try {
            await Excel.run(async context => {
                const config = await getConfig(context);
                if (config.has(ConfigKeys.AnswerLg)) {
                    this.setState({code: String(config.get(ConfigKeys.AnswerLg))});
                }
            });
        } catch (error) {
            this.props.addDebug(error);
        }
    };

    componentDidMount = () => {
        this.updateWidth();
        window.addEventListener('resize', this.updateWidth);
    };

    componentWillUnmount = () => {
        window.addEventListener('resize', this.updateWidth);
    };

    editorWillMount = (monaco: typeof monacoEditor) => {
        registerLGLanguage(monaco);
    };

    onChange = async (newValue, _e) => {
        this.setState({code: newValue});
    };

    clickTest = async () => {
        this.props.disableButton(this.buttonTest, true);
        this.props.pushEvent(Event.TestAnswerLg, this.state.code);
        setTimeout(() => this.props.clickDoSync(), 1000);
    };

    clickSyncLg = async () => {
        this.props.disableButton(this.buttonSyncLg, true);
        this.props.pushEvent(Event.SetAnswerLg, this.state.code);
        setTimeout(() => this.props.clickDoSync(), 1000);
    };

    clickSaveToExcel = async () => {
        this.props.disableButton(this.buttonSaveToExcel, true);

        try {
            await Excel.run(async context => {
                const configSheet = context.workbook.worksheets.getFirst();
                const configRange = configSheet.getUsedRange();
                configRange.load("values");
                configRange.load("columnIndex");
                configRange.load("rowIndex");

                await context.sync();

                let targetIndex = -1;
                for (let i = 0;i < configRange.values.length;++i) {
                    let element = configRange.values[i];
                    if (element.length < 2) continue;
                    if (String(element[0]).toLowerCase() == ConfigKeys.AnswerLg) {
                        targetIndex = i;
                        break;
                    }
                }
                if (targetIndex >= 0) {
                    const targetRange = configSheet.getRangeByIndexes(configRange.rowIndex + targetIndex, configRange.columnIndex + 1, 1, 1);
                    targetRange.values = [[this.state.code]];
                } else {
                    const targetRange = configSheet.getRangeByIndexes(configRange.rowIndex + configRange.values.length, configRange.columnIndex, 1, 2);
                    targetRange.values = [[ConfigKeys.AnswerLg, this.state.code]];
                }

                await context.sync();

                this.props.disableButton(this.buttonSaveToExcel, false);
            });
        } catch (error) {
            this.props.addDebug(error);
        }
    };

    render() {
        const {isOpened, height, width, code} = this.state;

        return (
            <div>
                <div>
                    <div style={{width: "100px", float:"left"}}>
                        <label>Lg Editor</label>
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
                        <button id={this.buttonTest} onClick={this.clickTest}>Test</button>
                        <button id={this.buttonSaveToExcel} onClick={this.clickSaveToExcel}>Save to Excel</button>
                        <button id={this.buttonSyncLg} onClick={this.clickSyncLg}>Sync LG</button>
                    </div>
                    <div ref={this.myRef} style={{height}}>
                        <MonacoEditor
                            width={width}
                            height={height}
                            theme={'lgtheme'}
                            language={'botbuilderlg'}
                            onChange={this.onChange}
                            value={code}
                            editorWillMount={this.editorWillMount}
                            />
                    </div>
                </Collapse>
            </div>
        );
      }
}
