import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./field-copy');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FIELD_COPY, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;

  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FIELD_COPY), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  // it('correctly handles a random API error', (done) => {
  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     return Promise.reject('An error has occurred');
  //   });

  //   cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', fromField: 'Title', toField: 'SSSS', listTitle: 'MyList' } }, (err?: any) => {
  //     try {
  //       assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //   });
  // });


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
    const actual = (command.validate() as CommandValidate)({ options: { debug: true, listTitle: 'MyList', fromField: 'Title', toField: 'SSSS' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'site.com', listTitle: 'MyList', fromField: 'Title', toField: 'SSSS' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the from field is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'MyList', toField: 'SSSS' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the to field is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'MyList', fromField: 'Title' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the listTitle is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', toField: 'ssss', fromField: 'Title' } });
    assert.notEqual(actual, true);
  });


  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', fromField: 'Title', toField: 'SSSS', listTitle: 'MyList' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.FIELD_COPY));
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

  it('Works wit text  to text  mapping', (done) => {
    let fetchedToFieldDef = false;
    let fetchedFromFieldDef = false;
    var fetchedRecords = false;
    sinon.stub(request, 'get').callsFake((opts) => {


      if (opts.url == `https://contoso.sharepoint.com/sites/sales/_api/web/lists/getByTitle('MyList')/fields?$filter=InternalName eq 'field1'`) {

        fetchedToFieldDef = true;
        console.log("Fetched to field");
        return Promise.resolve(
          { "value": [{ "InternalName": "field1", "TypeAsString": "Text" }] }
        );
      }
      if (opts.url == `https://contoso.sharepoint.com/sites/sales/_api/web/lists/getByTitle('MyList')/fields?$filter=InternalName eq 'field2'`) {
        fetchedFromFieldDef = true;
        console.log("Fetched from field");
        return Promise.resolve(
          // all we care about is the fieldtypekind and the internalname
          { "value": [{ "InternalName": "field2", "TypeAsString": "Text" }] }
          
        );
      }
      if (opts.url == `https://contoso.sharepoint.com/sites/sales/_api/web/lists/getByTitle('MyList')/items?$select=Id,xx,personId,person/Title&$expand=person&$orderBy=Id&$filter=Id gt 0&$top=50`) {
        fetchedRecords = true;
        console.log("Fetched records");
        return Promise.resolve(
          // empty for now
          { "value": [] }
        );
      }
      console.log(`url ${opts.url} did not match anything. Missing a test`);
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', fromField: 'field1', toField: 'field2', listTitle: 'MyList' } }, () => {
      try {

        console.log(`after actiion fetchedFromFieldDef is  ${fetchedFromFieldDef} fetchedToFieldDef is ${fetchedToFieldDef} fetchedRecords is ${fetchedRecords} the other is ${cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE'))} `);
        assert(fetchedFromFieldDef && fetchedToFieldDef && fetchedRecords);
        //assert(fetchedFromFieldDef && fetchedToFieldDef);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    Utils.restore(vorpal.find);//?
  });


});