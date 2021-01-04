import * as path from "path";

import { EmailMessageBuilder, ClientBuilder, FindItemsResultBuilder, FolderEnum } from "../lib";

const service = new ClientBuilder()
  .build();

const send = async () => {
  await new EmailMessageBuilder()
    .withService(service)
    .withSubject('Dentist Appointment')
    .withBody('The appointment is with Dr. Smith.')
    .withTo(['test@user.mail'])
    .withFileAttachment(path.join(__dirname, 'ordersWorkflowResult.txt'))
    .execute();
};

const readInbox = async () => {
  const items = await new FindItemsResultBuilder()
    .withService(service)
    .withFolder(FolderEnum.Inbox)
    .withTakeN(10)
    .execute();

  for (const item in items) {
    console.log(JSON.stringify(item, undefined, 2))
  }
};

const main = async () => {
  await send();
  console.log('Send Done!');

  await readInbox();
  console.log('Read Done!');
}

main();
