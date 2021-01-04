import { ExchangeService } from "ews-javascript-api";

import { ClientBuilder } from "./client";

describe('Client::', () => {
  test(`should allow to create ews client using builder`, async () => {
    const service = new ClientBuilder().build();

    expect(service)
      .toBeDefined();
    expect(service instanceof ExchangeService)
      .toBeTruthy();
  });
});
