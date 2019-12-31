import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
//import { number } from 'easy-table';
//import { ExceptionData } from 'applicationinsights/out/Declarations/Contracts';


const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listTitle: string;
  fromField: string;
  toField: string;
  setMissingValuesToBlank?: string;
  batchSize?: number;
  //fromList (copy based on lookup field)
  //filter (filter which items should be copied)
}

class SpoFieldCopyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.FIELD_COPY}`;
  }

  public get description(): string {
    return 'Copies the values from one field to another in the selected list';
  }

  public async commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): Promise<void> {

    let contextInfo: ContextInfo = await this.getRequestDigest(args.options.webUrl);
    // get the field defs

    let toFieldDef: any = await this.fetchFieldDefinition(args.options.webUrl, args.options.listTitle, args.options.toField, contextInfo, cmd, cb);
    if (!toFieldDef) { throw new Error("error fetching to field definition."); }
    let fromFieldDef: any = await this.fetchFieldDefinition(args.options.webUrl, args.options.listTitle, args.options.fromField, contextInfo, cmd, cb);
    if (!fromFieldDef) { throw new Error("error fetching from field definition."); }

    let lastId: number = 0;
    while (true) {
      let results = await this.getABatch(args, lastId, contextInfo, fromFieldDef.value[0], toFieldDef.value[0])
        .catch((err) => {
          cmd.log(vorpal.chalk.green('Error fetching batch'));
          this.handleRejectedODataJsonPromise(err, cmd, cb);

        });
      console.log(`@line 90 results.value is ${results.value}`)
      if (results.value.length === 0) break;
      if (args.options.batchSize && args.options.batchSize > 0) {
        await this.updateBatch(args, contextInfo.FormDigestValue, results.value, fromFieldDef.value[0], toFieldDef.value[0])
          .catch((err) => {
            cmd.log(vorpal.chalk.green('Error updating bacth'));
            this.handleRejectedODataJsonPromise(err, cmd, cb);

          });
      } else {
        await this.updateNoBatch(args, contextInfo.FormDigestValue, results.value, fromFieldDef.value[0], toFieldDef.value[0])
          .catch((err) => {
            cmd.log(vorpal.chalk.red('Error updating item'));
            this.handleRejectedODataJsonPromise(err, cmd, cb);

          });
      }

      lastId = results.value[results.value.length - 1].Id;
    }
    cmd.log(vorpal.chalk.green('DONE'));
    cb();

  }
  private async fetchFieldDefinition(webUrl: string, listTitle: string, fieldInternalName: string, contextInfo: ContextInfo, cmd: CommandInstance, cb: () => void): Promise<any> {
    const fieldDefinitionRequest: any = {
      url: `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/fields?$filter=InternalName eq '${fieldInternalName}'`,
      headers: {
        'X-RequestDigest': contextInfo.FormDigestValue,
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };
    let fromFieldDef: any = await request.get(fieldDefinitionRequest)
      .catch((err) => {
        cmd.log(vorpal.chalk.green('Error fetching from field'));
        this.handleRejectedODataJsonPromise(err, cmd, cb);
      });
    if (!fromFieldDef) {
      throw new Error("error fetching from field definition.");
    }
    return fromFieldDef;
  }

  private async updateNoBatch(args: any, formDigestValue: string, records: Array<any>, fromFieldDef: any, toFieldDef: any) {
    const listTitle = args.options.listTitle;
    //const toField = args.options.toField;
    //const fromField = args.options.fromField;

    const webUrl = args.options.webUrl;
    for (var record of records) {
      const body = await this.createUpdateJSON(args, toFieldDef, record, fromFieldDef, formDigestValue);
      //   console.log(body);
      const updateOptions: any = {
        url: `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${record.Id})`,
        headers: {
          'X-RequestDigest': formDigestValue,
          'Content-Type': `application/json;odata=verbose`,
          'Accept': `application/json;odata=verbose`,
          'Content-Length': body.length,
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*'

        },
        body: body
      }

      await request.post(updateOptions).catch((e) => {
        console.log(e);
      })

    }

  }


  private async updateBatch(args: any, formDigestValue: string, records: Array<any>, fromFieldDef: any, toFieldDef: any) {
    const listTitle = args.options.listTitle;
    // const toField = args.options.toField;
    // const fromField = args.options.fromField;
    const webUrl = args.options.webUrl;
    // see  https://social.technet.microsoft.com/wiki/contents/articles/30044.sharepoint-online-performing-batch-operations-using-rest-api.aspx
    var batchGuid = this.generateUUID();
    var batchContents = new Array();
    var changeSetId = this.generateUUID();
    batchContents.push('Content-Type: application/http');
    batchContents.push('Accept: application/json;odata=verbose');
    for (var record of records) {
      const body = await this.createUpdateJSON(args, toFieldDef, record, fromFieldDef, formDigestValue);
      var endpoint = `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${record.Id})`;
      batchContents.push('--changeset_' + changeSetId);
      batchContents.push('Content-Type: application/http');
      batchContents.push('Content-Transfer-Encoding: binary');
      batchContents.push('');
      batchContents.push('PATCH ' + endpoint + ' HTTP/1.1');
      batchContents.push('Content-Type: application/json;odata=verbose');
      batchContents.push('IF-MATCH:*');
      batchContents.push('');
      batchContents.push(`${body}`);
      batchContents.push('');
    }
    batchContents.push('--changeset_' + changeSetId + '--');



    var batchBody = batchContents.join('\u000d\u000a');
    batchContents = new Array();
    batchContents.push('--batch_' + batchGuid);
    batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
    batchContents.push('Content-Length: ' + batchBody.length);
    batchContents.push('Content-Transfer-Encoding: binary');
    batchContents.push('');
    batchContents.push(batchBody);
    batchContents.push('');

    const updateOptions: any = {
      url: `${webUrl}/_api/$batch`,
      headers: {
        'X-RequestDigest': formDigestValue,
        'Content-Type': `multipart/mixed; boundary="batch_${batchGuid}"`
      },
      body: batchContents.join('\r\n')
    }
    await request.post(updateOptions);
  }


  private GetItemTypeForListName(name: string) {
    return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
  }
  private generateUUID() {
    var d = new Date().getTime();
    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = (d + Math.random() * 16) % 16 | 0;
      d = Math.floor(d / 16);
      return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
    });
    return uuid;
  }
  private async createUpdateJSON(args: any, toFieldDef: any, record: any, fromFieldDef: any, formDigestValue: string): Promise<string> {
    // format update js0n based on from / to field types

    var update: any = {
      __metadata: {
        type: this.GetItemTypeForListName(args.options.listTitle)
      }
    };
    switch (fromFieldDef.TypeAsString + "|" + toFieldDef.TypeAsString) {
      case "Text|User":

        update[`${toFieldDef.InternalName}Id`] = await this.getNumericUserId(args, formDigestValue, record[fromFieldDef.InternalName]);
        //    console.log(update);
        break;
      case "Text|Text":
        update[`${toFieldDef.InternalName}`] = record[fromFieldDef.InternalName];
        break;
      case "Lookup|Text":
        // note that by default we take whatever field the lookup column refers to. We could take any field from yje other list.
        if (record[fromFieldDef.InternalName]) {
          update[`${toFieldDef.InternalName}`] = record[`${fromFieldDef.InternalName}`][`${fromFieldDef.LookupField}`];
        } else {
          if (args.options.setMissingValuesToBlank === 'true') {
            update[`${toFieldDef.InternalName}`] = "";
          } else {
          }

        }
        break;
      default:
        throw new Error(`No mapping defined to convert ${fromFieldDef.TypeAsString} to  ${toFieldDef.TypeAsString}. Have you considered contributing? `)
    }

    const body = JSON.stringify(update);
    //   console.log("body");
    //  console.log(body);
    // return the OBJECT not the string!!!
    return body;
  }

  private async getNumericUserId(args: any, formDigestValue: string, userEmail: string): Promise<number | null> {

    var logonName = `i:0#.f|membership|${userEmail}`;
    const ensureUserOption: any = {
      url: `${args.options.webUrl}/_api/web/ensureuser`,
      headers: {
        'X-RequestDigest': formDigestValue,
        'Content-Type': `application/json;odata=verbose`,
        'Accept': `application/json`,
      },
      body: JSON.stringify({ 'logonName': logonName })
    }
    var id: number | null = null;
    await request.post(ensureUserOption)
      .then((userresult: any) => {
        userresult = JSON.parse(userresult);
        //     console.log(userresult.Id);
        id = userresult.Id;
      })
      .catch((err) => {
        console.log(`user ${userEmail} was not found`);
        id = null;
      });

    //   console.log(`id is ${id}`)
    return id;

  };
  private async getABatch(args: any, lastId: number, contextInfo: ContextInfo, fromFieldDef: any, toFieldDef: any) {
    const listTitle = args.options.listTitle;
    const webUrl = args.options.webUrl;
    const batchSize = args.options.batchSize ? args.options.batchSize : 50;
    var requestUrl = "";
    //  console.log("fromFieldDef in getabatch")
    //  console.log(fromFieldDef);
    //  console.log(fromFieldDef.TypeAsString);
    //  console.log(fromFieldDef["TypeAsString"]);

    switch (fromFieldDef.TypeAsString + "|" + toFieldDef.TypeAsString) {
      case "Lookup|Text":
        requestUrl = `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$select=Id,${fromFieldDef.InternalName}Id,${toFieldDef.InternalName},${fromFieldDef.InternalName}/${fromFieldDef.LookupField}&$expand=${fromFieldDef.InternalName}&$orderBy=Id&$filter=Id gt ${lastId}&$top=${batchSize}`;
        break;
      case "Text|Text":
        requestUrl = `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$select=Id,${fromFieldDef.InternalName},${toFieldDef.InternalName}&$orderBy=Id&$filter=Id gt ${lastId}&$top=${batchSize}`;
        break;
      case "Text|User":
        requestUrl = `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$select=Id,${fromFieldDef.InternalName},${toFieldDef.InternalName}Id,${toFieldDef.InternalName}/Title&$expand=${toFieldDef.InternalName}&$orderBy=Id&$filter=Id gt ${lastId}&$top=${batchSize}`;
        break;
      default:
        throw new Error(`No mapping defined to convert ${fromFieldDef.TypeAsString} to  ${toFieldDef.TypeAsString}. Have you considered contributing? `)
    }
    //  console.log(`Fetching data using url ${requestUrl}`);
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'X-RequestDigest': contextInfo.FormDigestValue,
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };
    // console.log(requestOptions.url);
    let results: any = await request.get(requestOptions);
    return results;
  }



  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site containing the list or library  where the field should be copied'
      },
      {
        option: '-f, --fromField <fromField>',
        description: 'The intenal name of the field to copy from'
      },
      {
        option: '-t, --toField <toField>',
        description: 'The intenal name of the field to copy to'
      },
      {
        option: '-bs, --batchSize <batchSize>',
        description: 'The number of rows to process at once'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the field should be copied'
      },

      {
        option: '--setMissingValuesToBlank <setMissingValuesToBlank>',
        description: 'If the source field is blank or empty, set the destination to blank.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {


    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter url missing';
      }
      if (!args.options.listTitle) {
        return 'Required parameter list missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.fromField) {
        return 'Required parameter fromfield missing';
      }
      if (!args.options.toField) {
        return 'Required parameter toField missing';
      }

      // // get the field defs
      // let contextInfo: ContextInfo = await this.getRequestDigest(args.options.webUrl);
      // const requestOptions: any = {
      //   url: `${args.options.webUrl}/_api/web/lists/getByTitle('${args.options.listTitle}')/fields?$filter=InternalName in ('${args.options.fromField}','${args.options.toField}')}`,
      //   headers: {
      //     'X-RequestDigest': contextInfo.FormDigestValue,
      //     accept: 'application/json;odata=nometadata'
      //   },
      //   json: true
      // };
      // let results: any = await request.get(requestOptions);
      // return results;


      // if (args.options.options) {
      //   let optionsError: string | boolean = true;
      //   const options: string[] = ['SetMissingValuesToBlank'];
      //   args.options.options.split(',').forEach(o => {
      //     o = o.trim();
      //     if (options.indexOf(o) < 0) {
      //       optionsError = `${o} is not a valid value for the options argument. Allowed values are SetMissingValuesToBlank`;
      //     }
      //   });
      //   return optionsError;
      // }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    This command copies the values in a list from one field to another. Both fields must already exist in the list.
    If either field does not exists you will get an error
    ${chalk.grey('field not found????.')} error.

  Examples:
  
    copyies the the value of one text column to another text column (useful if you want to change the internal name of a column)
    o365 spo field copy -u https://russellwgove.sharepoint.com/sites/clitest --listTitle test  -f Lookup_x0020_Column -t yy --batchSize 100 --setMissingValuesToBlank true
  

    copyies the text value of a lookup column to a text column
    o365 spo field copy -u https://russellwgove.sharepoint.com/sites/clitest --listTitle test  -f Lookup_x0020_Column -t yy --batchSize 100 --setMissingValuesToBlank true

    takes the value of the source field, looks up a user with that value, and sets the user in the result field.
    (this is useful if you have imprted a list with an email address and you want to convert that to a user field! )
    o365 spo field copy -u https://russellwgove.sharepoint.com/sites/clitest --listTitle test  -f xx -t person --batchSize 1
    
  More information:

    AddFieldOptions enumeration
      https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.addfieldoptions.aspx
`);
  }
}

module.exports = new SpoFieldCopyCommand();