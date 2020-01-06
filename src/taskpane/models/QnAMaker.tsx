export class QnADTO {
    answer: string;
    source: string;
    questions: string[];

    constructor(answer: string, source: string, question: string){
        this.answer = answer;
        this.source = source;
        this.questions = [ question ];
    }
}
