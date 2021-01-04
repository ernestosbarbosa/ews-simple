import { ExchangeService } from "ews-javascript-api";

import { ClientBuilder } from "./client";
import { EmailMessageBuilder } from "./message";

describe('Message::', () => {
  let service: ExchangeService;

  beforeAll(() => {
    service = new ClientBuilder()
      .withUser(`<EWS username>`)
      .withPwd(`<EWS user pwd>`)
      .build();
  });

  test(`should throw if service wasn't initialized`, async () => {
    await expect(new EmailMessageBuilder()
      .withSubject('Dentist Appointment')
      .withBody('The appointment is with Dr. Smith.')
      .withTo(['yelebe4156@liaphoto.com'])
      .execute())
      .rejects.toThrow(new Error(`Ews service wasn't initialized!`));
  });

  test(`should allow to send email using builder`, async () => {
    await new EmailMessageBuilder()
      .withService(service)
      .withSubject('Dentist Appointment')
      .withBody('The appointment is with Dr. Smith.')
      .withTo(['yelebe4156@liaphoto.com'])
      .execute();
  });
});
