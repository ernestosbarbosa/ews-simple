import { ExchangeService } from "ews-javascript-api";

import { ClientBuilder } from "./client";
import { FindItemsResultBuilder } from "./item.results";

describe('Item Results::', () => {
  let service: ExchangeService;

  beforeAll(() => {
    service = new ClientBuilder()
      .withUser(`<EWS username>`)
      .withPwd(`<EWS user pwd>`)
      .build();
  });

  test(`should allow to read emails from the specified folder`, async () => {
    const items = await new FindItemsResultBuilder()
      .withService(service)
      .withQueryString('isread:false')
      .withGetRaw(true)
      .withSaveAttachments(true)
      .withMarkAsRead(false)
      .execute();

    expect(items)
      .toBeDefined();
  });

  test(`should allow to read emails and save attachments`, async () => {
    const items = await new FindItemsResultBuilder()
      .withService(service)
      .withQueryString('isread:false')
      .withMarkAsRead(false)
      .withSaveAttachments(true)
      .execute();

    expect(items)
      .toBeDefined();
  });

});
