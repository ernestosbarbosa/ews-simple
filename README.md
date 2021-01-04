# ews-simple

## Introduction
This package is build on top of **ews-javascript-api** and makes performing EWS operations easier.

## API
### ClientBuilder
This builder should allow you to configure EWS service with the specified url and auth data.

Supported methods:
  1. **withVersion** *(value: ews.ExchangeVersion)* - configures ews vesrsion
  2. **withUser** *(value: string)* - configures ews user name
  3. **withPwd** *(value: string)* - configures user pwd
  4. **withUri** *(value: string)* - configures ews server uri
  5. **withToken** *(value: string)* - configures user authorization token

#### Example
```ts
const service = new ClientBuilder()
  .withUser(`<EWS username>`)
  .withPwd(`<EWS user pwd>`)
  .build();
```

### FindItemsResultBuilder
This builder should allow you to return search results using the specified input folder and query string.

Supported methods:
  1. **withService** *(value: ews.ExchangeService)* - allow set ews service to be linked to
  2. **withFolder** *(value: FolderEnum)* - allow to set ews folder to be looked into
  3. **withTakeN** *(value: number)* - allow to set the amount(default *1000*) of items to be returned
  4. **withOffset** *(value: number)* - allow to set offset
  5. **withOffsetBasePoint** *(value: ews.OffsetBasePoint)* - allow to set base point(*start* or *end*)
  6. **withQueryString** *(value: string)* - allow to set query string for search
  7. **withMarkAsRead** *(value: boolean)* - allow to set whether to mark letter as read or not
  8. **withDelete** *(value: boolean)* - allow to set whether do delete letter after read
  9. **withGetRaw** *(value: boolean)* - allow to set wherther to return raw items or parse them
  10. **withSaveAttachments** *(value: boolean)* - allow to set whether to download attachments or not
  11. **withDownloadsFolder** *(value: string)* - allow to set downloads folder for attachments(default is *downloads*)
  12. **withRawBody** *(value: boolean)*- allow to set whether to return raw body or parse it before return

#### Examples

#### Read (get Subject, Body and Received Date) 10 unread mails from Inbox
```ts
const items = await new FindItemsResultBuilder()
  .withService(service)
  .withFolder(FolderEnum.Inbox)
  .withQueryString('isread:false')
  .withTakeN(10)
  .execute();
```

#### Read (keep them raw) 10 read mails from Favorites
```ts
const items = await new FindItemsResultBuilder()
  .withService(service)
  .withFolder(FolderEnum.Favorites)
  .withQueryString('isread:true')
  .withGetRaw(true)
  .execute();
```

### EmailMessageBuilder
This builder should allow you to send a email with the specified data.

  1. **withService** *(value: ews.ExchangeService)* - allow set ews service to be linked to
  2. **withSubject** *(value: string)* - allow to set message subject
  3. **withBodyType** *(value: ews.BodyType)* - allow to set message body content type
  4. **withBody** *(value: string)* - allow to set message body content
  5. **withTo** *(value: string[])* - allow to set a list of **To** recepients
  6. **withСс** *(value: string[])* - allow to set a list of **Сс** recepients
  7. **withFileAttachment** *(value: string)* - allow to set path to file attachment

#### Examples
##### Send email
```ts
await new EmailMessageBuilder()
  .withService(service)
  .withSubject(<mail subject>)
  .withBodyType(ews.BodyType.HTML)
  .withBody(<mail body>)
  .withTo(<recepients list>)
  .execute();
```

##### Send email with file attachment
```ts
await new EmailMessageBuilder()
  .withService(service)
  .withSubject(<mail subject>)
  .withBodyType(ews.BodyType.Text)
  .withBody(<mail body>)
  .withTo(<recepients list>)
  .withFileAttachment(<file name>)
  .execute();
```