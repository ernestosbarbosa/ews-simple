import * as ews from 'ews-javascript-api';
export declare class EmailMessageBuilder {
    _service: ews.ExchangeService | undefined;
    _subject: string;
    _bodyType: ews.BodyType;
    _body: string;
    _to: string[] | undefined;
    _attachment: string | undefined;
    withService(value: ews.ExchangeService): this;
    withSubject(value: string): this;
    withBodyType(value: ews.BodyType): this;
    withBody(value: string): this;
    withTo(value: string[]): this;
    withFileAttachment(value: string): this;
    execute(): Promise<void>;
}
