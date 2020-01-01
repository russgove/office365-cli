import { IFieldTransformer,IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export  class textToTextFieldTransformer implements IFieldTransformer {
  setQuery(fieldInternalName: string): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fieldInternalName];
    var expands: Array<string> = [];
    return { selects: selects, expands: expands }
  }
 async setJSON(listitem: any,fromFieldDef:IFieldDefinition,toFieldDef:IFieldDefinition,transformerDefinition:ITransformerDefinition): Promise<any> {
    let update:any={};
    update[`${toFieldDef.InternalName}`]=listitem[fromFieldDef.InternalName]
    return update;
  }

}
