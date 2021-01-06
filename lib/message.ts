/*
 * EWS EmailMessage builder
 *
 * 
 */

// Dependencies
import * as ews from 'ews-javascript-api';
import * as fs from 'fs';
import * as util from 'util';

const debug = util.debuglog('message');

export class EmailMessageBuilder {
  _service: ews.ExchangeService | undefined;
  _subject = 'Dentist Appointment';
  _bodyType = ews.BodyType.Text;
  _body = 'The appointment is with Dr. Smith.';
  _to: string[] | undefined;
  _cc: string[] | undefined;
  _attachment: string | undefined;

  withService(value: ews.ExchangeService) {
    this._service = value;
    return this;
  }

  withSubject(value: string) {
    this._subject = value;
    return this;
  }

  withBodyType(value: 'text' | 'html') {
    this._bodyType = value === 'text' ? ews.BodyType.Text : ews.BodyType.HTML;
    return this;
  }

  withBody(value: string) {
    this._body = value;
    return this;
  }

  withTo(value: string[]) {
    this._to = value;
    return this;
  }

  withCc(value: string[]) {
    this._cc = value;
    return this;
  }

  withFileAttachment(value: string) {
    this._attachment = value;
    return this;
  }

  async execute() {
    if (this._service === undefined) {
      throw new Error(`Ews service wasn't initialized!`);
    }
    if (this._to === undefined) {
      throw new Error(`Recepients list wasn't initialized!`);
    }
    debug('Prepare message object...');
    const message = new ews.EmailMessage(this._service);
    message.Subject = this._subject;
    message.Body = new ews.MessageBody(this._bodyType, this._body);
    this._to.map(user => message.ToRecipients.Add(user));
    if (this._cc) {
      this._cc.map(user => message.CcRecipients.Add(user));
    }

    if (this._attachment) {
      try {
        debug('Adding attachment...');
        /*
         * AddFileAttachment parameters - filename (name of attachment as shown in outlook/owa not actual file from disk),
         * base64 content of file (read file from disk and convert to base64 yourself)
         */
        const content = fs.readFileSync(this._attachment, 'utf8');
        message.Attachments.AddFileAttachment(this._attachment, Buffer.from(content).toString('base64'));
        debug('Sending mail and saving a copy in sent folder...');
        await message.SendAndSaveCopy();
        debug('Message was successfully sent!');
      } catch (err) {
        console.error(err);
      }
    }
  }
}
