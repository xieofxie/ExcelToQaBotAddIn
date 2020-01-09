import { QnAMakerModels } from "@azure/cognitiveservices-qnamaker";

export class Event {
    // QnA
    static GetQnA: string = "GetQnA";
    static EnableQnA: string = "EnableQnA";
    static CreateQnA: string = "CreateQnA";
    static AddQnA: string = "AddQnA";
    static DelQnA: string = "DelQnA";
    static UpdateQnA: string = "UpdateQnA";
    // Source
    static AddSource: string = "AddSource";
    static DelSource: string = "DelSource";
    // Configs
    static SetMinScore: string = "SetMinScore";
    static SetResultNumber: string = "SetResultNumber";
    // Answer Lg
    static SetAnswerLg: string = "SetAnswerLg";
    static TestAnswerLg: string = "TestAnswerLg";
    // Others
    static SetDebug: string = "SetDebug";
}

export class SourceType {
    static Editorial = "Editorial";
    static File = "File";
    static Url = "Url";
}

export class Source {
    id: string;
    description: string;
    type: string;

    constructor(Id: string = '', Description: string = '', Type: string = '') {
        this.id = Id;
        this.description = Description;
        this.type = Type;
    }
}

export interface AddSourceEvent extends QnAMakerModels.UpdateKbOperationDTOAdd {
    knowledgeBaseId: string;
    qnaListId?: string;
    qnaListDescription?: string;
    urlsDescription?: string[];
    filesDescription?: string[];
}

export class DelSourceEvent {
    knowledgeBaseId: string;
    ids: string[];
}

// TODO use official pacakge
export class QnAMakerEndpoint {
    knowledgeBaseId: string;
    endpointKey: string;
    host: string;

    constructor(id: string = '', key: string = '', host: string = '') {
        this.knowledgeBaseId = id;
        this.endpointKey = key;
        this.host = host;
    }
}
export class QnAMakerEndpointEx extends QnAMakerEndpoint {
    name: string;
    enable: boolean;
    // Map does not serialize
    sources: { [index: string]: Source; };
}
