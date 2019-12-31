import { IFieldTransformer } from "./IFieldTransformer";
export  class textToTextFieldTransformer implements IFieldTransformer {
  setQuery(fieldInternalName: string): { selects: Array<string>; expands: Array<string> } {
    var selects: Array<string> = [fieldInternalName];
    var expands: Array<string> = [];
    return { selects: selects, expands: expands }
  }
  setJSON(listitem: any,fieldInternalName:string): any {
    return { fieldInternalName: listitem[fieldInternalName]}
  }

}
