import * as ews from 'ews-javascript-api';
export declare class EmailMessageBuilder {
    _service: ews.ExchangeService | undefined;
    _subject: string;
    _bodyType: ews.BodyType;
    _body: string;
    _to: string[] | undefined;
    _cc: string[] | undefined;
    _attachment: string | undefined;
    withService(value: ews.ExchangeService): this;
    withSubject(value: string): this;
    withBodyType(value: 'text' | 'html'): this;
    withBody(value: string): this;
    withTo(value: string[]): this;
    withCc(value: string[]): this;
    withFileAttachment(value: string): this;
    execute(): Promise<void>;
}
