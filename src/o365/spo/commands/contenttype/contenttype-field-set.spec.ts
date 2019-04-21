import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./contenttype-field-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import config from '../../../../config';

describe(commands.CONTENTTYPE_FIELD_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
    (command as any).requestDigest = '';
    (command as any).siteId = '';
    (command as any).webId = '';
    (command as any).fieldLink = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.CONTENTTYPE_FIELD_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.CONTENTTYPE_FIELD_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a field reference to content type updating field schema', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return Promise.resolve();
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`);
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a field reference to content type updating field schema (debug)', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return Promise.resolve();
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`);
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
        }

        return Promise.reject(`[
  {
    "SchemaVersion": "15.0.0.0",
    "LibraryVersion": "16.0.7911.1206",
    "ErrorInfo": {
      "ErrorMessage": "Invalid request",
      "ErrorValue": null,
      "TraceCorrelationId": "59577d9e-70af-0000-22fb-870cf639feff",
      "ErrorCode": -2130575252,
      "ErrorTypeName": "InvalidRequest"
    },
    "TraceCorrelationId": "59577d9e-70af-0000-22fb-870cf639feff"
  }
]`)
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a field reference to content type updating field schema from AllowDeletion=FALSE', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"FALSE\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return Promise.resolve();
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`);
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a field reference to content type without updating field schema', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`);
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds a field reference to content type without updating field schema (debug)', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`);
        }

        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7911.1206",
              "ErrorInfo": null,
              "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
            }
          ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while updating field schema', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          SchemaXml: "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        })) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while adding field reference', (done) => {
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
  {
    "SchemaVersion": "15.0.0.0",
    "LibraryVersion": "16.0.7911.1206",
    "ErrorInfo": {
      "ErrorMessage": "An error has occurred",
      "ErrorValue": null,
      "TraceCorrelationId": "1e5a7d9e-9047-0000-22fb-8361c9a5b96e",
      "ErrorCode": -2130575252,
      "ErrorTypeName": "Error"
    },
    "TraceCorrelationId": "1e5a7d9e-9047-0000-22fb-8361c9a5b96e"
  }
]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates existing field link', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates existing field link (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates existing field link (hidden)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'false', hidden: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates existing field link (required)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
              {
                "SchemaVersion": "15.0.0.0",
                "LibraryVersion": "16.0.7911.1206",
                "ErrorInfo": null,
                "TraceCorrelationId": "73557d9e-007f-0000-22fb-89971360c85c"
              }
            ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while trying to retrieve field link', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          'odata.null': true
        });
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:field:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": null,
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    },
    5,
    {
      "IsNull": false
    },
    6,
    {
      "_ObjectIdentity_": "e5547d9e-705d-0000-22fb-8faca5696ed8|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b"
    }
  ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find field link for field 5ee2dd25-d941-455a-9bdb-7f2c54aed11b`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips updating when existing field link is up-to-date (no values specified)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips updating when existing field link is up-to-date (no values specified; debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error while updating the field link', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        return Promise.resolve({
          "FieldInternalName": null,
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Name": "PnPAlertStartDateTime",
          "Required": false
        });
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\" AllowDeletion=\"TRUE\"><Default>[today]</Default></Field>"
        });
      }

      if (opts.url.indexOf(`_api/site?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "50720268-eff5-48e0-835e-de588b007927"
        });
      }

      if (opts.url.indexOf(`_api/web?$select=Id`) > -1) {
        return Promise.resolve({
          "Id": "d1b7a30d-7c22-4c54-a686-f1c298ced3c7"
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F:fl:5ee2dd25-d941-455a-9bdb-7f2c54aed11b" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:50720268-eff5-48e0-835e-de588b007927:web:d1b7a30d-7c22-4c54-a686-f1c298ced3c7:contenttype:0x0100558D85B7216F6A489A499DB361E1AE2F" /></ObjectPaths></Request>`) {
          return Promise.resolve(`[
    {
      "SchemaVersion": "15.0.0.0",
      "LibraryVersion": "16.0.7911.1206",
      "ErrorInfo": {
        "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
      },
      "TraceCorrelationId": "e5547d9e-705d-0000-22fb-8faca5696ed8"
    }
  ]`);
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Unknown Error`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if site URL is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'site.com', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the content type ID is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the field ID is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified field ID is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } });
    assert.equal(actual, true);
  });

  it('fails validation if the specified required value is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified hidden value is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', hidden: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the required option is set to true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation when the required option is set to false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'false' } });
    assert.equal(actual, true);
  });

  it('passes validation when the hidden option is set to true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', hidden: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation when the hidden option is set to false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', hidden: 'false' } });
    assert.equal(actual, true);
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures contentTypeId as string option', () => {
    const types = (command.types() as CommandTypes);
    ['contentTypeId', 'c'].forEach(o => {
      assert.notEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.CONTENTTYPE_FIELD_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('handles error while retrieving request digest', (done) => {
    Utils.restore((command as any).getRequestDigestForSite);
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } }));
    let fieldLinksRequestNum: number = 0;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')/fieldlinks('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')`) > -1) {
        fieldLinksRequestNum++;
        switch (fieldLinksRequestNum) {
          case 1:
            return Promise.resolve({
              'odata.null': true
            });
          case 2:
            return Promise.resolve({
              "FieldInternalName": null,
              "Hidden": false,
              "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
              "Name": "PnPAlertStartDateTime",
              "Required": false
            });
        }
      }

      if (opts.url.indexOf(`_api/web/fields('5ee2dd25-d941-455a-9bdb-7f2c54aed11b')?$select=SchemaXmlWithResourceTokens`) > -1) {
        return Promise.resolve({
          "SchemaXmlWithResourceTokens": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"4\"><Default>[today]</Default></Field>"
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', contentTypeId: '0x0100558D85B7216F6A489A499DB361E1AE2F', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b', required: 'true', hidden: 'true' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', contentTypeId: '0x0100FF0B2E33A3718B46A3909298D240FD93', fieldId: '5ee2dd25-d941-455a-9bdb-7f2c54aed11b' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});