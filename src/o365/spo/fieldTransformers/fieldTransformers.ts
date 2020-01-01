import {textToTextFieldTransformer} from "./textToTextFieldTransformer";
export interface IFieldTransformer {

  /**
   * @param fieldInternalName the internal name of the field to be transformed.
   * 
   * @returns  an array of fields to be added to the selects clause of the request and an array of expands to be 
   * to the $expands clause of the request
   */
  setQuery(fieldInternalName: string,transformerDefinition:ITransformerDefinition): { selects: Array<string>; expands: Array<string> };

  /**
     * @param listitem : The listitem selected from the sharepoint list that will include the fields requested in the 
   * selects and expands returned from setqury.
   * 
   * @param fieldInternalName :The internalName of the field in the listitem used to create the result
   * 
   * @returns an object that can be used to update the target field type
   */
   setJSON(listitem: any, fromFieldDef: IFieldDefinition,toFieldDefinition:IFieldDefinition,transformationDefinition:ITransformerDefinition,webUrl:string,formDigestValue:string): Promise<any>;
}
export interface IFieldDefinition {
  InternalName: string;
  TypeAsString: string;
  // add other attribute here as needed (like Prson Or group, DateOnly, etc,,)
}
export interface ITransformerDefinition{
  fromFieldType:string,
   toFieldType:string,
   name:string,
   transformer: IFieldTransformer,
   description:string,
   // add other swithches here and pass to the transformer. Thatway a single transformer can be reuesed by passing different switches
}

var transfomers:Array<ITransformerDefinition>=[
     //// add other swithches her and pass to the transformer. Thatway a single transformer can be reuesed by passing different switches (like replace nonEmpty values, system Update, etc)
  {fromFieldType:"Text", toFieldType:"Text",transformer:new textToTextFieldTransformer(),name:"TextToText",description:"Text to Text-- can be used to change the internal name of a field"}
]
export default transfomers;