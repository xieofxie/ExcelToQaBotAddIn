export class Event {
    static SetQnA: string = "SetQnA";
    static SetResultNumber: string = "SetResultNumber";
    static SetMinScore: string = "SetMinScore";
    static SetAnswerLg: string = "SetAnswerLg";
    static SetDebug: string = "SetDebug";
    static TestAnswerLg: string = "TestAnswerLg";
}

export class QnAMakerEndpoint {
    KnowledgeBaseId: string;
    EndpointKey: string;
    Host: string;

    constructor(id = null, key = null, host = null) {
        this.KnowledgeBaseId = id;
        this.EndpointKey = key;
        this.Host = host;
    }
}
