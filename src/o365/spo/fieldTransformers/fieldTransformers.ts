import {textToTextFieldTransformer} from "./textToTextFieldTransformer";
import {IFieldTransformer} from './IFieldTransformer';
export interface ITransformerDefinition{
  fromFieldType:string,
   toFieldType:string,
   name:string,
   transformer: IFieldTransformer,
   description:string
}

var transfomers:Array<ITransformerDefinition>=[
  {fromFieldType:"Text", toFieldType:"Text",transformer:new textToTextFieldTransformer(),name:"text to tyext:",description:"Text to Text"}
]
export default transfomers;