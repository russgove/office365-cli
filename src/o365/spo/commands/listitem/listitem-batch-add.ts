import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
//import { ListItemInstance } from './ListItemInstance';
import { FolderExtensions } from '../../FolderExtensions';
import * as path from 'path';
import { Transform } from 'stream';
const vorpal: Vorpal = require('../../../../vorpal-init');
const csv = require('@fast-csv/parse');
import {createReadStream} from 'fs';


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  contentType?: string;
  folder?: string;
  path: string;
}

class SpoListItemAddCommand extends SpoCommand {

  public allowUnknownOptions(): boolean | undefined {
    return false;
  }

  public get name(): string {
    return commands.LISTITEM_BATCH_ADD;
  }

  public get description(): string {
    return 'Creates a list item in the specified list for each row in the specified .csv file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.contentType = typeof args.options.contentType !== 'undefined';
    telemetryProps.folder = typeof args.options.folder !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let lineNumber: number = 0;
    let contentTypeName: string | null = null;
    let listRestUrl: string | null = null;
    let batchSize: number = 4000      ; // max is  1048576
    let recordsToAdd = "";
    let batchInProcess = false;
    const  generateUUID= function() {
      var d = new Date().getTime();
      var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c == 'x' ? r : (r & 0x7 | 0x8)).toString(16);
      });
      return uuid;
    }
    const parseResults = (response: any): void => {
      cmd.log(`in parseresults`)
      let responseLines = response.toString().split('\n');
      // read each line until you find JSON...
      for (let responseLine of responseLines) {
       try {
          // parse the JSON response...
          var tryParseJson = JSON.parse(responseLine);
          for (let result of tryParseJson.d.AddValidateUpdateItemUsingPath.results){
            if (result.HasException){
              cmd.log(result)
            }
          }
        } catch (e) {
        }
      }

    }

    let targetFolderServerRelativeUrl: string = ``;
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = Utils.getSafeFileName(path.basename(fullPath));
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    listRestUrl = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    const folderExtensions: FolderExtensions = new FolderExtensions(cmd, this.debug);

    if (this.verbose) {
      cmd.log(`Getting content types for list...`);
    }

    const requestOptions: any = {
      url: `${listRestUrl}/contenttypes?$select=Name,Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((response: any): Promise<void> => {
        if (args.options.contentType) {
          const foundContentType = response.value.filter((ct: any) => {
            const contentTypeMatch: boolean = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;

            if (this.debug) {
              cmd.log(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
            }

            return contentTypeMatch;
          });

          if (this.debug) {
            cmd.log('content type filter output...');
            cmd.log(foundContentType);
          }

          if (foundContentType.length > 0) {
            contentTypeName = foundContentType[0].Name;
          }

          // After checking for content types, throw an error if the name is blank
          if (!contentTypeName || contentTypeName === '') {
            return Promise.reject(`Specified content type '${args.options.contentType}' doesn't exist on the target list`);
          }

          if (this.debug) {
            cmd.log(`using content type name: ${contentTypeName}`);
          }
        }

        if (args.options.folder) {
          if (this.debug) {
            cmd.log('setting up folder lookup response ...');
          }

          const requestOptions: any = {
            url: `${listRestUrl}/rootFolder`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          }

          return request
            .get<any>(requestOptions)
            .then(rootFolderResponse => {
              targetFolderServerRelativeUrl = Utils.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], args.options.folder as string);

              return folderExtensions.ensureFolder(args.options.webUrl, targetFolderServerRelativeUrl);
            });
        }
        else {
          return Promise.resolve();
        }
      })
      .then((): any => {
        if (this.verbose) {
          cmd.log(`Creating items in list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}...`);
        }
       
        let lmapRequestBody = this.mapRequestBody;
        //start the batch
        let changeSetId = generateUUID();
        let endpoint = `${listRestUrl}/AddValidateUpdateItemUsingPath()`;
         // get the file
         let fileStream = createReadStream(fileName);
        let csvStream: any = csv.parseStream(fileStream, {headers:true  })
        csvStream
        .pipe(new Transform({
          objectMode: true,
          write(row: any, encoding: string, callback: (error?: (Error | null)) => void): void {
            console.log(`Processing row ${JSON.stringify(row)}`)
          
              console.log(`Done processing row ${JSON.stringify(row)}`);
              lineNumber++;
              const requestBody: any = {
                formValues: lmapRequestBody(row)
              };
              if (args.options.folder) {
                requestBody.listItemCreateInfo = {
                  FolderPath: {
                    DecodedUrl: targetFolderServerRelativeUrl
                  }
                };
              }
              if (args.options.contentType && contentTypeName !== '') {
                requestBody.formValues.push({
                  FieldName: 'ContentType',
                  FieldValue: contentTypeName
                });
              }
              // row is ready
              recordsToAdd += '--changeset_' + changeSetId + '\u000d\u000a' +
                'Content-Type: application/http' + '\u000d\u000a' +
                'Content-Transfer-Encoding: binary' + '\u000d\u000a' +
                '\u000d\u000a' +
                'POST ' + endpoint + ' HTTP/1.1' + '\u000d\u000a' +
                'Content-Type: application/json;odata=verbose' + '\u000d\u000a' +
                'Accept: application/json;odata=verbose' + '\u000d\u000a' +
                '\u000d\u000a' +
                `${JSON.stringify(requestBody)}` + '\u000d\u000a' +
                '\u000d\u000a';
  
              if (recordsToAdd.length >= batchSize) {
                /***  Send the batch   **/
                recordsToAdd += '--changeset_' + changeSetId + '--' + '\u000d\u000a';
                csvStream.pause();
                batchInProcess = true;
                let batchContents = new Array();
                let batchId = generateUUID();
                batchContents.push('--batch_' + batchId);
                batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
                batchContents.push('Content-Length: ' + recordsToAdd.length);
                batchContents.push('Content-Transfer-Encoding: binary');
                batchContents.push('');
                batchContents.push(recordsToAdd);
                batchContents.push('');
                const updateOptions: any = {
                  url: `${args.options.webUrl}/_api/$batch`,
                  headers: {
                    'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
                  },
                  body: batchContents.join('\r\n')
                }
                request.post(updateOptions)
                  .catch((e) => {
                    cb(e);
                  })
                  .then((response) => {
                    parseResults(response)
                  })
                  .finally(() => {
                    recordsToAdd = ``;
                    changeSetId = generateUUID();
                    csvStream.resume();
                    batchInProcess = false;
                  })
              }
              this.push(row);
              callback();
      
          },
        }))
          .on("headers",function(headers:any){ // not gerring called. should validate column names up front
             cmd.log("in headers")
             cmd.log(headers)
            csvStream.pause();
            cmd.log(headers);
            csvStream.resume();
          }) 
          .on("end", function () {
            function waitForLastBatch() { // should not need to do this. "end" should not be called while stream is paused
              if (batchInProcess === true) {
                setTimeout(waitForLastBatch, 5000);
              } else {
                if (recordsToAdd.length > 0) {
                  let batchContents = new Array();
                  let batchId = generateUUID();
                  batchContents.push('--batch_' + batchId);
                  batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
                  batchContents.push('Content-Length: ' + recordsToAdd.length);
                  batchContents.push('Content-Transfer-Encoding: binary');
                  batchContents.push('');
                  batchContents.push(recordsToAdd);
                  batchContents.push('');
                  const updateOptions: any = {
                    url: `${args.options.webUrl}/_api/$batch`,
                    headers: {
                      'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
                    },
                    body: batchContents.join('\r\n')
                  }
                  request.post(updateOptions)
                    .catch((e) => {
                      cb(e);
                    })
                    .then((response) => {
                      parseResults(response)
                    })
                    .finally(() => {
                      cmd.log(`Processed ${lineNumber} Rows`)
                      cb();
                    })
                } else {
                  cmd.log(`Processed ${lineNumber} Rows`)
                  cb();
                }
              }
            }
            waitForLastBatch();
          })
          .on("error", function (error: any) {
            cb(error)
          });
      })
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the item should be added'
      },
      {
        option: '-p, --path <path>',
        description: 'the path of the csv file with records to be added to the SharePoint list'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list where the item should be added. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list where the item should be added. Specify listId or listTitle but not both'
      },
      {
        option: '-c, --contentType [contentType]',
        description: 'The name or the ID of the content type to associate with the new item'
      },
      {
        option: '-f, --folder [folder]',
        description: 'The list-relative URL of the folder where the item should be created'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }



  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.listId && !args.options.listTitle) {
        return `Specify listId or listTitle`;
      }

      if (!args.options.path) {
        return `Specify path`;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Specify listId or listTitle but not both`;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Add an item with Title ${chalk.grey('Demo Item')} and content type name ${chalk.grey('Item')} to list with
    title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --contentType Item --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Item"

    Add an item with Title ${chalk.grey('Demo Multi Managed Metadata Field')} and
    a single-select metadata field named ${chalk.grey('SingleMetadataField')} to list with
    title ${chalk.grey('Demo List')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Managed Metadata Field" --SingleMetadataField "TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;"

    Add an item with Title ${chalk.grey('Demo Multi Managed Metadata Field')} and a multi-select
    metadata field named ${chalk.grey('MultiMetadataField')} to list with title ${chalk.grey('Demo List')}
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Managed Metadata Field" --MultiMetadataField "TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;"
  
    Add an item with Title ${chalk.grey('Demo Single Person Field')} and a single-select people
    field named ${chalk.grey('SinglePeopleField')} to list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Single Person Field" --SinglePeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'}]"
      
    Add an item with Title ${chalk.grey('Demo Multi Person Field')} and a multi-select people
    field named ${chalk.grey('MultiPeopleField')} to list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Multi Person Field" --MultiPeopleField "[{'Key':'i:0#.f|membership|markh@conotoso.com'},{'Key':'i:0#.f|membership|adamb@conotoso.com'}]"
    
    Add an item with Title ${chalk.grey('Demo Hyperlink Field')} and a hyperlink field named
    ${chalk.grey('CustomHyperlink')} to list with title ${chalk.grey('Demo List')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LISTITEM_ADD} --listTitle "Demo List" --webUrl https://contoso.sharepoint.com/sites/project-x --Title "Demo Hyperlink Field" --CustomHyperlink "https://www.bing.com, Bing"
   `);
  }

  private mapRequestBody(row: any): any {
    const requestBody: any = [];
    Object.keys(row).forEach(async key => {
      requestBody.push({ FieldName: key, FieldValue: (<any>row)[key] });
    });
    return requestBody;
  }
}

module.exports = new SpoListItemAddCommand();