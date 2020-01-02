/**
 * To add a new transformer you need to
 * 1. add a class to the field transformers foler that implements the IFieldTransformer interface
 * 2. import that class here.
 * 3. add a new element to the transfoRmers array at the bottom of this file
 * 
 */

import {textToTextFieldTransformer} from "./textToTextFieldTransformer";
import {textToUserFieldTransformer} from "./textToUserFieldTransformer";
import {lookupToTextTransformer} from "./LookupToTextFieldTransformer";

export interface IFieldTransformer {

  /**
   * @param fieldInternalName the internal name of the field to be transformed.
   * 
   * @returns  an array of fields to be added to the selects clause of the request and an array of expands to be 
   * to the $expands clause of the request
   */
  setQuery(fromFieldDef: IFieldDefinition,transformerDefinition:ITransformerDefinition): { selects: Array<string>; expands: Array<string> };

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
  LookupField: string;
  LookupList: string;
  LookupWebId: string;
  // add other attribute here as needed (like Prson Or group, DateOnly, etc,,)
}
export interface ITransformerDefinition{
  fromFieldType:string, // TypeAsString of the field we are copying From
   toFieldType:string, // TypeAsString of the field we are copying  to
   name:string, // the name of the transformer (anyting will do)
   transformer: IFieldTransformer, // the implementatopm of the transformer
   description:string, // a descriptopm of the transformation
   searchDisplayName?:boolean;// used in text to User, Search using display namme? If no, then use email
   // add other swithches here and pass to the transformer. Thatway a single transformer can be reuesed by passing different switches
}

var transfomers:Array<ITransformerDefinition>=[
   //// add other swithches her and pass to the transformer. Thatway a single transformer can be reuesed by passing different switches (like replace nonEmpty values, system Update, etc) (or maybe use a closure)
  {fromFieldType:"Text", toFieldType:"Text",transformer:new textToTextFieldTransformer(),name:"TextToText",description:"Text to Text-- can be used to change the internal name of a field"},
  {fromFieldType:"Text", toFieldType:"User",transformer:new textToUserFieldTransformer(),name:"EmailTextToUser",description:"Can be used to convert a column containing a persons email to a User Column. (Single user only)",searchDisplayName:false},
  {fromFieldType:"Text", toFieldType:"User",transformer:new textToUserFieldTransformer(),name:"DisplayNameTextToUser",description:"Can be used to convert a column containing a persons email to a User Column. (Single user only)",searchDisplayName:true},
  {fromFieldType:"Lookup", toFieldType:"Text",transformer:new lookupToTextTransformer(),name:"LookupToTextDefault",description:"Can be used to copy the default lookup Value of a lookup column to a text field"}

]
export default transfomers;