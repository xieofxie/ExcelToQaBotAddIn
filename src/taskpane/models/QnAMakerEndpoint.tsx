export default class QnAMakerEndpoint {
    KnowledgeBaseId;
    EndpointKey;
    Host;

    constructor(id = null, key = null, host = null) {
        this.KnowledgeBaseId = id;
        this.EndpointKey = key;
        this.Host = host;
    }
}
