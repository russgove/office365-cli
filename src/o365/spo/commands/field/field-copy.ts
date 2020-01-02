import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import transformers, { IFieldDefinition, ITransformerDefinition } from '../../fieldTransformers/fieldTransformers';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listTitle: string;
  fromField: string;
  toField: string;
  transformer: string;
  batchSize?: number;
  whatIf?: boolean;
  filter?: string;
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

    let fromFieldDef: IFieldDefinition = await this.fetchFieldDefinition(args.options.webUrl, args.options.listTitle, args.options.fromField, contextInfo, cmd, cb);
    if (!fromFieldDef) { throw new Error("error fetching from field definition."); }
    let toFieldDef: IFieldDefinition = await this.fetchFieldDefinition(args.options.webUrl, args.options.listTitle, args.options.toField, contextInfo, cmd, cb);
    if (!toFieldDef) { throw new Error("error fetching to field definition."); }

    let transformerDefinition: ITransformerDefinition | null = null;
    for (let transformer of transformers) {
      if (transformer.fromFieldType === fromFieldDef.TypeAsString
        && transformer.toFieldType === toFieldDef.TypeAsString
        && transformer.name === args.options.transformer) {
        transformerDefinition = transformer;
        break;
      }
    }
    if (!transformerDefinition) {

      cmd.log(`No transformer named ${args.options.transformer} for converting from ${fromFieldDef.TypeAsString} to ${toFieldDef.TypeAsString} could be found. Valid transformers follow:`)
      for (let transformer of transformers) {
        if (transformer.fromFieldType === fromFieldDef.TypeAsString
          && transformer.toFieldType === toFieldDef.TypeAsString
        ) {
          cmd.log(`${transformer.name}  (${transformer.description})`);
        }
      }
      cb();
    } else {
      let sande = transformerDefinition.transformer.setQuery(fromFieldDef, transformerDefinition);
      let selects = sande.selects;
      let expands = sande.expands;
      let fetchQuery: string = this.createFetchQuery(args, selects, expands);
      let lastId: number = 0;
      while (true) {
        let results = await this.getABatch(fetchQuery, lastId, contextInfo)
          .catch((err) => {
            cmd.log(vorpal.chalk.green('Error fetching batch'));
            this.handleRejectedODataJsonPromise(err, cmd, cb);

          });
        if (results.value.length === 0) break;
        if (args.options.batchSize && args.options.batchSize > 0) {
          await this.updateBatch(args, contextInfo.FormDigestValue, results.value, transformerDefinition, fromFieldDef, toFieldDef)
            .catch((err) => {
              cmd.log(vorpal.chalk.green('Error updating bacth'));
              this.handleRejectedODataJsonPromise(err, cmd, cb);

            });
        } else {
          await this.updateNoBatch(args, contextInfo.FormDigestValue, results.value, transformerDefinition, fromFieldDef, toFieldDef)
            .catch((err) => {
              cmd.log(vorpal.chalk.red('Error updating item'));
              this.handleRejectedODataJsonPromise(err, cmd, cb);

            });
        }
        lastId = results.value[results.value.length - 1].Id;
      }
    }



    cmd.log(vorpal.chalk.green('DONE'));
    cb();

  }
  private async fetchFieldDefinition(webUrl: string, listTitle: string, fieldInternalName: string, contextInfo: ContextInfo, cmd: CommandInstance, cb: () => void): Promise<IFieldDefinition> {
    const fieldDefinitionRequest: any = {
      url: `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/fields?$filter=InternalName eq '${fieldInternalName}'`,
      headers: {
        'X-RequestDigest': contextInfo.FormDigestValue,
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };
    return request.get<{ value: Array<IFieldDefinition> }>(fieldDefinitionRequest)
      .then((results: { value: Array<IFieldDefinition> }) => {

        if (results.value.length !== 0) {
          return results.value[0];
        }
        else {
          console.log(`query for field ${fieldInternalName} returned no results`);
          throw new Error(`Field ${fieldInternalName} was not found`);
        }
      })
      .catch((err) => {
        throw new Error(`Error fetching fiield definiton for  ${fieldInternalName}`);
      });

  }

  private async updateNoBatch(args: any, formDigestValue: string, records: Array<any>, transformer: ITransformerDefinition, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition) {
    const listTitle = args.options.listTitle;
    //const toField = args.options.toField;
    //const fromField = args.options.fromField;

    const webUrl = args.options.webUrl;
    for (let record of records) {
      let body = await this.createUpdateJSON(args, fromFieldDef, toFieldDef, record, transformer, formDigestValue);
      let postBody = JSON.stringify(body);
      const updateOptions: any = {
        url: `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${record.Id})`,
        headers: {
          'X-RequestDigest': formDigestValue,
          'Content-Type': `application/json;odata=verbose`,
          'Accept': `application/json;odata=verbose`,
          'Content-Length': postBody.length,
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*'

        },
        body: postBody
      }

      await request.post(updateOptions).catch((e) => {
        throw (e)
      })

    }

  }
  private async updateBatch(args: any, formDigestValue: string, records: Array<any>, transformer: ITransformerDefinition, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition) {
    const listTitle = args.options.listTitle;
    // const toField = args.options.toField;
    // const fromField = args.options.fromField;
    const webUrl = args.options.webUrl;
    // see  https://social.technet.microsoft.com/wiki/contents/articles/30044.sharepoint-online-performing-batch-operations-using-rest-api.aspx
    let batchGuid = this.generateUUID();
    let batchContents = new Array();
    let changeSetId = this.generateUUID();
    batchContents.push('Content-Type: application/http');
    batchContents.push('Accept: application/json;odata=verbose');
    for (let record of records) {
      const reqbody = await this.createUpdateJSON(args, fromFieldDef, toFieldDef, record, transformer, formDigestValue);
      let postBody = JSON.stringify(reqbody);
      let endpoint = `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${record.Id})`;
      batchContents.push('--changeset_' + changeSetId);
      batchContents.push('Content-Type: application/http');
      batchContents.push('Content-Transfer-Encoding: binary');
      batchContents.push('');
      batchContents.push('PATCH ' + endpoint + ' HTTP/1.1');
      batchContents.push('Content-Type: application/json;odata=verbose');
      batchContents.push('IF-MATCH:*');
      batchContents.push('');
      batchContents.push(`${postBody}`);
      batchContents.push('');
    }
    batchContents.push('--changeset_' + changeSetId + '--');



    let batchBody = batchContents.join('\u000d\u000a');
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
    let d = new Date().getTime();
    let uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      let r = (d + Math.random() * 16) % 16 | 0;
      d = Math.floor(d / 16);
      return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
    });
    return uuid;
  }


  private async createUpdateJSON(args: any, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition, record: any, transformerDefinitiom: ITransformerDefinition, formDigestValue: string): Promise<string> {
    // format update json based on from / to field types
    let update: any = await transformerDefinitiom.transformer.setJSON(record, fromFieldDef, toFieldDef, transformerDefinitiom, args.options.webUrl, formDigestValue);
    update["__metadata"] = {
      type: this.GetItemTypeForListName(args.options.listTitle)
    };
    return update;
  }

  // private async getNumericUserId(args: any, formDigestValue: string, userEmail: string): Promise<number | null> {

  //   let logonName = `i:0#.f|membership|${userEmail}`;
  //   const ensureUserOption: any = {
  //     url: `${args.options.webUrl}/_api/web/ensureuser`,
  //     headers: {
  //       'X-RequestDigest': formDigestValue,
  //       'Content-Type': `application/json;odata=verbose`,
  //       'Accept': `application/json`,
  //     },
  //     body: JSON.stringify({ 'logonName': logonName })
  //   }
  //   let id: number | null = null;
  //   await request.post(ensureUserOption)
  //     .then((userresult: any) => {
  //       userresult = JSON.parse(userresult);
  //       //     console.log(userresult.Id);
  //       id = userresult.Id;
  //     })
  //     .catch((err) => {
  //       console.log(`user ${userEmail} was not found`);
  //       id = null;
  //     });

  //   //   console.log(`id is ${id}`)
  //   return id;

  // };
  private createFetchQuery(args: any, selects: Array<string>, expands: Array<string>): string {
    const listTitle = args.options.listTitle;
    const webUrl = args.options.webUrl;
    const batchSize = args.options.batchSize ? args.options.batchSize : 50;
    let requestUrl = "";

    let effectiveSelects: Array<string> = ["Id", ...selects];
    requestUrl = `${webUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$select=${effectiveSelects.join(',')}`;
    if (expands && expands.length > 0) {
      const x = expands.join(',');
      requestUrl += `&$expand=${x}`;
    }
    //{{{lastId}}} gets replaced for each batch
    
    if (args.options.filter){
      requestUrl += `&$filter=${args.options.filter} and Id gt {{{lastId}}}`;
    }else{
      requestUrl += `&$filter=Id gt {{{lastId}}}`;
    }

    requestUrl += `&$orderBy=Id&$top=${batchSize}`;
    return requestUrl;


  }


  private async getABatch(requestUrl: string, lastId: number, contextInfo: ContextInfo): Promise<any> {
    const effectiveRequestUrl = requestUrl.replace(`{{{lastId}}}`, lastId.toString());
    const requestOptions: any = {
      url: effectiveRequestUrl,
      headers: {
        'X-RequestDigest': contextInfo.FormDigestValue,
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };
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
        description: 'The number of rows to process at once . Specifying batchsize of 0 bypasses batching logic and processes records one at a'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the field should be copied'
      },

      {
        option: '--setMissingValuesToBlank <setMissingValuesToBlank>',
        description: 'If the source field is blank or empty, set the destination to blank. Keep switch or make separate transformers'
      },
      {
        option: '--transformer <transformer>',
        description: 'The name of the field transformer to use (should we default this somehow, perhaps for like field types)'
      },
      {
        option: '-w --whatif <whatif>',
        description: 'If the whatif flag is set to try, the command will display the updates that would be made, but not actually make any updates'
      },
      {
        option: '--filter <filter>',
        description: 'An odata filter to be added to the request to select the listitems to update'
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
      if (!args.options.transformer) {
        return 'Required parameter transformer missing';
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