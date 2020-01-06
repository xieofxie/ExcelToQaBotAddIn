export class Event {
    // QnA
    static GetQnA: string = "GetQnA";
    static EnableQnA: string = "EnableQnA";
    static CreateQnA: string = "CreateQnA";
    static AddQnA: string = "AddQnA";
    static DelQnA: string = "DelQnA";
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

// TODO use official pacakge
export class QnADTO {
    answer: string;
    questions: string[];

    constructor(answer: string, question: string){
        this.answer = answer;
        this.questions = [ question ];
    }
}

export class Source {
    Id: string;
    Description: string;
    Type: string;

    constructor(Id: string = null, Description: string = null, Type: string = null) {
        this.Id = Id;
        this.Description = Description;
        this.Type = Type;
    }
}

export class SourceEvent extends Source {
    KnowledgeBaseId: string;
    QnaList: QnADTO[];
}

// TODO use official pacakge
export class QnAMakerEndpoint {
    knowledgeBaseId: string;
    endpointKey: string;
    host: string;

    constructor(id: string = null, key: string = null, host: string = null) {
        this.knowledgeBaseId = id;
        this.endpointKey = key;
        this.host = host;
    }
}

// TODO special lower case
export class QnAMakerEndpointEx extends QnAMakerEndpoint {
    name: string;
    enable: boolean;
    // Map does not serialize
    sources: { [index: string]: Source; };
}

export class EnableQnAEvent {
    KnowledgeBaseId: string;
    Enable: boolean;

    constructor(KnowledgeBaseId: string = null, Enable: boolean = null) {
        this.KnowledgeBaseId = KnowledgeBaseId;
        this.Enable = Enable;
    }
}
