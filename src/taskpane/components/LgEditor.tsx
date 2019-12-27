import * as React from "react";
import MonacoEditor from 'react-monaco-editor';
import {Collapse} from 'react-collapse';
import * as monacoEditor from 'monaco-editor/esm/vs/editor/editor.api';
import { registerLGLanguage } from './lg';
import { throttle } from 'lodash'

export interface AppProps {
}

interface AppState {
    isOpened: boolean;
    height: number;
    width: number;
}

export class LgEditor extends React.Component<AppProps, AppState> {
    myRef = null;

    constructor(props: AppProps, context: any) {
        super(props, context);
        this.state = {
            isOpened: false,
            height: 500,
            width: 500
        };
        this.myRef = React.createRef();
    }

    updateWidth = throttle(() => {
        const width = this.myRef.current.getBoundingClientRect().width;
        if (width != this.state.width) {
            this.setState({width: width});
        }
    }, 100);

    componentDidMount() {
        this.updateWidth();
        window.addEventListener('resize', this.updateWidth);
    }

    componentWillUnmount() {
        window.addEventListener('resize', this.updateWidth);
    }

    editorWillMount = (monaco: typeof monacoEditor) => {
        registerLGLanguage(monaco);
    };

    render() {
        const {isOpened, height, width} = this.state;

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
                    <div ref={this.myRef} style={{height}}>
                        <MonacoEditor
                            width={width}
                            height={height}
                            theme={'lgtheme'}
                            language={'botbuilderlg'}
                            editorWillMount={this.editorWillMount}
                            />
                    </div>
                </Collapse>
            </div>
        );
      }
}
