
import { IFieldTransformer, IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export class lookupToTextTransformer implements IFieldTransformer {
  setQuery(fromFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fromFieldDef.InternalName + 'Id', fromFieldDef.InternalName + '/' + fromFieldDef.LookupField];
    var expands: Array<string> = [fromFieldDef.InternalName];
    return { selects: selects, expands: expands }
  }
  async setJSON(listitem: any, fromFieldDef: IFieldDefinition, toFieldDef: IFieldDefinition, transformerDefinition: ITransformerDefinition, webUrl: string, formDigestValue: string): Promise<any> {
    console.log(`in setJson o flookup to text  webUrl is ${webUrl}`);
    let update: any = {};
    update[`${toFieldDef.InternalName}`] = listitem[fromFieldDef.LookupField]; //haha just copies title
    console.log(`lookup field name is ${fromFieldDef.LookupField}  `)
    console.log(`item is ${JSON.stringify(listitem)}`);
    let value= {};
    if (listitem[fromFieldDef.InternalName]){
       value = listitem[fromFieldDef.InternalName][fromFieldDef.LookupField];
    }

    const path = `${fromFieldDef.InternalName}/${fromFieldDef.LookupField}`
    console.log(`path is  ${path} value is ${value}`)

    update[`${toFieldDef.InternalName}`] = value;

    return update;
  }


}
