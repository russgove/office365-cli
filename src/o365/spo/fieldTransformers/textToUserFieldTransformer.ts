import request from '../../../request';
import { IFieldTransformer,IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export  class textToUserFieldTransformer implements IFieldTransformer {
  setQuery(fromFieldDef: IFieldDefinition,transformerDefinition:ITransformerDefinition): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fromFieldDef.InternalName];
    var expands: Array<string> = [];
    return { selects: selects, expands: expands }
  }
  async setJSON(listitem: any,fromFieldDef:IFieldDefinition,toFieldDef:IFieldDefinition,transformerDefinition:ITransformerDefinition,webUrl: string,formDigestValue:string): Promise< any >{
  
    let update:any={};
    update[`${toFieldDef.InternalName}Id`] = await this.getNumericUserId(webUrl,formDigestValue, listitem[fromFieldDef.InternalName]);
    return update;
  }
  private async getNumericUserId(webUrl:string, formDigestValue: string, userEmail: string): Promise<number | null> {
    var logonName = `i:0#.f|membership|${userEmail}`;
    const ensureUserOption: any = {
      url: `${webUrl}/_api/web/ensureuser`,
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
        id = userresult.Id;
      })
      .catch((err) => {
        console.log(`user ${userEmail} was not found`);
        id = null;
      });
    return id;

  };

}
