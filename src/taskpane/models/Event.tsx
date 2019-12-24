export class Event {
    static SetQnA: string = "SetQnA";
    static SetResultNumber: string = "SetResultNumber";
    static SetMinScore: string = "SetMinScore";
    static SetNoResultResponse: string = "SetNoResultResponse";
    static SetDebug:string = "SetDebug";
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
