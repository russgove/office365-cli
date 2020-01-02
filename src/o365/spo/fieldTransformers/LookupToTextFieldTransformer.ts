
import { IFieldTransformer, IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export class lookupToTextTransformer implements IFieldTransformer {
  setQuery(fromFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fromFieldDef.InternalName + 'Id', fromFieldDef.InternalName + '/' + fromFieldDef.LookupField];
    var expands: Array<string> = [fromFieldDef.InternalName];
    return { selects: selects, expands: expands }
  }
  async setJSON(args:any,listitem: any, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition, webUrl: string, formDigestValue: string): Promise<any> {
    let update: any = {};
    let value= {};
    if (listitem[fromFieldDef.InternalName]){
       value = listitem[fromFieldDef.InternalName][fromFieldDef.LookupField];
    }
    update[`${toFieldDef.InternalName}`] = value;
    return update;
  }


}
