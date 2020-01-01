import { IFieldTransformer,IFieldDefinition, ITransformerDefinition } from "./fieldTransformers";
export  class textToPersonFieldTransformer implements IFieldTransformer {
  setQuery(fromInternalName: string,transformerDefinition:ITransformerDefinition): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fromInternalName];
    var expands: Array<string> = [];
    return { selects: selects, expands: expands }
  }
  async setJSON(listitem: any,fromFieldDef:IFieldDefinition,toFieldDef:IFieldDefinition,transformerDefinition:ITransformerDefinition): Promise< any >{
    let update:any={};
    update[`${toFieldDef.InternalName}`]=listitem[fromFieldDef.InternalName]
    return update;
  }

}
