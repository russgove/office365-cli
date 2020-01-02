import request from '../../../request';
import { IFieldTransformer, IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export class textToTextInOtherListFieldTransformer implements IFieldTransformer {
  setQuery(fromFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fromFieldDef.InternalName];
    var expands: Array<string> = []; // no joins here. This assumes ther is no relationship between the lists
    return { selects: selects, expands: expands }
  }
  async setJSON(args: any, listitem: any, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition, webUrl: string, formDigestValue: string): Promise<any> {
    let update: any = {};

    update[`${toFieldDef.InternalName}Id`] = await this.getValueFromOtherList(args, webUrl, formDigestValue, listitem[fromFieldDef.InternalName]);
  }
  async getValueFromOtherList(args: any, webUrl: string, formDigestValue: string, joinFieldValue: string): Promise<string| null> {
    const otherListRequest: any = {
      url: `${webUrl}/_api/web/lists/getByTitle('${args.otherListTitle}')/items?$filter='${args.otherListJoinFieldName}' eq '${joinFieldValue}'`,
      headers: {
        'X-RequestDigest': formDigestValue,
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };
    const result: string | null = await request.get(otherListRequest)
      .then((userresult: any) => {
        userresult = JSON.parse(userresult);
        if (userresult.value.length === 0) {
          console.log(`Querying list  '${args.otherListJoinFieldName}' where field '${args.otherListJoinFieldName}' equals '${joinFieldValue}' returned no values`);
          return null;
        }
        else {
          if (userresult.value.length > 1) {
            console.log(`Querying list  '${args.otherListJoinFieldName}' where field '${args.otherListJoinFieldName}' equals '${joinFieldValue}' returned ${userresult.value.length} values. Unsure which to use, so no action taken.`);
            return null;
          }
        }

        return userresult.value[0][`${args.otherListTargetFieldName}`];
      })
      .catch((err) => {
        console.log(err);

        console.log(`an arror occurred fetching item from  list  '${args.otherListJoinFieldName}' where field '${args.otherListJoinFieldName}' equals '${joinFieldValue}' , so no action taken.`);
        return null;
      });

    return result;


    return "ZZ";
  }


}
